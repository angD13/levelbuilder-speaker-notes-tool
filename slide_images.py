#!/usr/bin/env python3
"""
slide_images.py

Fetches per-slide thumbnail images straight from the Google Slides API,
so the Streamlit app can show the actual slide next to its extracted
speaker notes -- no need to keep a separate window open with the deck
just to cross-check the tool's work.

Reuses the SAME Google OAuth session as the Drive-API pptx download
(download_pptx_via_drive_api in speaker_notes_sync.py) -- one
credentials.json / token.json pair, one login, rather than a second
separate Google sign-in for a second feature.

The token now needs a BROADER scope than before (Drive export AND
Slides thumbnails). If a token.json saved under the old Drive-only
scope already exists, get_google_credentials() notices it's missing
the newly-required scope and transparently discards it, opening the
browser again for one fresh consent that covers both -- never a raw
"insufficient scope" API error surfaced to the user, and never more
than that one extra login.

IMPORTANT, confirmed with a real run: this does NOT hand the app a
thumbnail URL to load directly. getThumbnail's contentUrl is a direct
link to Google's own image-serving host, and handing several of those
to a browser at once (one per slide shown) gets some of them rate-
limited with a 429 -- confirmed by opening one of the failed image
URLs directly, which also 429'd (and it was a DIFFERENT slide failing
each reload, not the same one every time, which rules out it being
about any particular slide's content). So this module downloads each
thumbnail's bytes itself instead, one at a time with a small pacing
gap and a retry-with-backoff on 429, and hands the app the raw image
bytes -- which also means there's no URL left to expire later, so the
app doesn't need to track a TTL on anything it caches from here.
"""

import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/presentations.readonly",
]


def get_google_credentials(credentials_path: str, token_path: str, scopes: list = None):
    """
    Shared credentials loader for every Google API call this tool makes
    (Drive export, Slides thumbnails) -- one token file, kept in sync
    with whatever the UNION of scopes across all callers is (see
    GOOGLE_SCOPES above). Centralizing this means adding a new Google
    API call later only means adding its scope to GOOGLE_SCOPES; this
    function already handles the "existing token doesn't cover the new
    scope yet" migration.
    """
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow

    scopes = scopes or GOOGLE_SCOPES
    token_file = Path(token_path)

    creds = None
    if token_file.exists():
        creds = Credentials.from_authorized_user_file(str(token_file), scopes)
        # A token saved before a scope was added to GOOGLE_SCOPES is
        # missing it entirely -- Credentials.from_authorized_user_file
        # doesn't fail on this by itself (it just builds an object that
        # will get a "insufficient scope" error the moment it's used
        # for the new scope's API), so check for it explicitly and
        # treat it the same as "no saved session" rather than letting
        # that surface as a confusing API error later.
        if creds and creds.scopes and not set(scopes).issubset(set(creds.scopes)):
            creds = None

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None

    if not creds or not creds.valid:
        print("No valid saved session found -- opening a browser to log in...")
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
        creds = flow.run_local_server(port=0)
        token_file.write_text(creds.to_json())
        print(f"Login successful. Session saved to {token_path} for future runs.")

    return creds


def _is_rate_limited(exc: Exception) -> bool:
    """True for a 429 from either the Slides API (googleapiclient's
    HttpError, status on exc.resp.status) or the raw image download
    (urllib.error.HTTPError, status on exc.code) -- both surfaces can
    hit it, confirmed by a real 429 opening a thumbnail URL directly."""
    status = getattr(getattr(exc, "resp", None), "status", None)
    if status is None:
        status = getattr(exc, "code", None)
    return status == 429


def _with_retries(fn, max_attempts: int = 4):
    """Runs fn() (a zero-arg callable), retrying with exponential
    backoff (1s, 2s, 4s, ...) only on a 429 -- any other error raises
    immediately, same as before."""
    delay = 1.0
    for attempt in range(max_attempts):
        try:
            return fn()
        except Exception as e:
            if _is_rate_limited(e) and attempt < max_attempts - 1:
                time.sleep(delay)
                delay *= 2
                continue
            raise


def fetch_slide_thumbnails(
    presentation_id: str, credentials_path: str, token_path: str, slide_numbers: list = None
) -> dict:
    """
    Returns {slide_number (1-indexed): thumbnail_image_bytes} for every
    slide in the deck, or just the given `slide_numbers` when passed --
    cheaper when only a handful of slides are actually needed (e.g. the
    specific slides referenced by the plan, rather than the whole deck).

    Downloads each slide's thumbnail BYTES itself (one at a time, with
    a small pacing gap and a retry-with-backoff on 429) rather than
    handing the caller a live URL to load -- see the module docstring
    for why: Google's image-serving host rate-limits several of those
    URLs being fetched by a browser at once, and this sidesteps it
    entirely by doing the fetching here, one request at a time.
    """
    try:
        from googleapiclient.discovery import build
    except ImportError:
        sys.exit(
            "Missing dependencies for slide thumbnails. Run:\n"
            "  pip install google-api-python-client google-auth-oauthlib"
        )

    creds = get_google_credentials(credentials_path, token_path)
    service = build("slides", "v1", credentials=creds)

    presentation = service.presentations().get(
        presentationId=presentation_id, fields="slides.objectId"
    ).execute()
    slides = presentation.get("slides", [])
    wanted = set(slide_numbers) if slide_numbers else None

    thumbnails = {}
    for i, slide in enumerate(slides, start=1):
        if wanted is not None and i not in wanted:
            continue
        try:
            # NOTE: deliberately NOT passing thumbnailProperties.thumbnailSize
            # here -- googleapiclient's dynamically-generated method for this
            # endpoint rejects that dotted keyword outright ("Got an unexpected
            # keyword argument"), confirmed against a real call. Omitting it
            # just falls back to the API's own default size, which is still
            # fine since the app only ever displays these scaled down to
            # 140-160px anyway.
            thumb = _with_retries(lambda: service.presentations().pages().getThumbnail(
                presentationId=presentation_id,
                pageObjectId=slide["objectId"],
            ).execute())
            image_bytes = _with_retries(lambda: urllib.request.urlopen(thumb["contentUrl"], timeout=15).read())
            thumbnails[i] = image_bytes
        except Exception as e:
            # One slide failing (transient API hiccup, a slide type that
            # doesn't thumbnail cleanly, a 429 that outlasted the
            # retries, etc.) shouldn't lose every OTHER slide in the
            # same batch -- confirmed this was happening before: a
            # single bad slide raised out of the loop entirely, so
            # cache.update() in app.py never ran at all for that whole
            # batch. Skip it and keep going; app.py already falls back
            # to a plain caption when a slide has no image.
            print(f"Warning: couldn't fetch thumbnail for slide {i}: {e}", file=sys.stderr)
        # A small gap between slides, independent of the above retry
        # backoff -- spreads out the burst of requests to Google's
        # image-serving host instead of firing them back-to-back,
        # which is what triggered the 429s in the first place.
        time.sleep(0.3)

    return thumbnails
