import threading
import time
from googleapiclient.errors import HttpError

# Single semaphore shared across PerformanceDataUtils and TrackerDataUtils.
# Caps total concurrent Google API calls regardless of how many checks run in parallel.
_api_semaphore = threading.Semaphore(3)


def execute_with_retry(request, max_attempts=6):
    """Execute a Google API request with exponential backoff on 429/5xx errors.

    max_attempts=6 gives backoffs of 2, 4, 8, 16, 32 s (62 s total) before giving up,
    which outlasts a 60-second quota window after a burst.
    """
    for attempt in range(max_attempts):
        try:
            with _api_semaphore:
                return request.execute()
        except HttpError as e:
            if e.resp.status in (429, 500, 502, 503, 504) and attempt < max_attempts - 1:
                delay = 2 ** (attempt + 1)
                print(f"API error {e.resp.status}, retrying in {delay}s (attempt {attempt + 1}/{max_attempts})...")
                time.sleep(delay)
            else:
                raise
