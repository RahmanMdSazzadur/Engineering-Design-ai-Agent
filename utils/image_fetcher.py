"""
image_fetcher.py — Download a representative image for a machine using
DuckDuckGo Images (no API key required).

Usage
-----
    from utils.image_fetcher import fetch_machine_image

    image_path = fetch_machine_image("Siemens SIMOTICS 1LE1 15 kW Induction Motor")
    # Returns a Path to a temp file, or None if no image was found.
"""

from __future__ import annotations

import logging
import tempfile
import urllib.request
from pathlib import Path

logger = logging.getLogger(__name__)

# Common image extensions we'll accept
_VALID_EXTS = {"jpg", "jpeg", "png", "webp"}


def fetch_machine_image(machine_name: str) -> Path | None:
    """Search DuckDuckGo Images for *machine_name* and return a local temp file path.

    Parameters
    ----------
    machine_name:
        Human-readable name of the machine (e.g. ``"Kuka KR 10 R1100 robot"``).

    Returns
    -------
    Path or None
        Path to a downloaded temp image file, or ``None`` if no image could be
        fetched (network unavailable, package missing, or no results).
    """
    try:
        from ddgs import DDGS  # new package name
    except ImportError:
        try:
            from duckduckgo_search import DDGS  # old package name fallback
        except ImportError:
            logger.warning(
                "ddgs package not installed — skipping image fetch. "
                "Install it with: pip install ddgs"
            )
            return None

    query = f"{machine_name} machine product"
    logger.info("Searching DuckDuckGo Images for: %r", query)

    try:
        with DDGS() as ddgs:
            results = list(ddgs.images(query, max_results=8))
    except Exception as exc:
        logger.warning("DuckDuckGo image search failed: %s", exc)
        return None

    for result in results:
        url: str = result.get("image") or result.get("thumbnail") or ""
        if not url:
            continue

        # Determine extension from URL (strip query string first)
        raw_ext = url.split("?")[0].rsplit(".", 1)[-1].lower()
        ext = raw_ext if raw_ext in _VALID_EXTS else "jpg"

        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "Mozilla/5.0 (compatible; DatasheetAgent/1.0)"},
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = resp.read()

            if len(data) < 1_000:  # skip suspiciously tiny files
                continue

            tmp = tempfile.NamedTemporaryFile(suffix=f".{ext}", delete=False)
            tmp.write(data)
            tmp.close()
            logger.info("Machine image saved to %s (from %s)", tmp.name, url)
            return Path(tmp.name)

        except Exception as exc:
            logger.debug("Could not download %s: %s", url, exc)
            continue

    logger.warning("No usable image found for %r", machine_name)
    return None
