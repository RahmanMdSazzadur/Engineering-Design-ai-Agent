"""
web_searcher.py — Search the web for a machine's real product page and
datasheet URLs before the LLM extraction call.

The results are injected into the Gemini prompt as verified context so
the model uses real URLs instead of hallucinating them.
"""

from __future__ import annotations

import logging
from typing import Optional

logger = logging.getLogger(__name__)


def search_machine_web_context(machine_name: str) -> Optional[str]:
    """Search DuckDuckGo for the machine's official product page and datasheet.

    Returns a formatted string of top search results (title + URL + snippet)
    ready to be injected into the LLM prompt as verified web context.
    Returns None if the search fails or the package is unavailable.
    """
    try:
        try:
            from ddgs import DDGS
        except ImportError:
            from duckduckgo_search import DDGS  # fallback to old name

    except ImportError:
        logger.warning("ddgs not installed — skipping web search context")
        return None

    queries = [
        f"{machine_name} official product page site specifications",
        f"{machine_name} datasheet PDF manufacturer",
    ]

    all_results: list[str] = []

    try:
        with DDGS() as ddgs:
            for query in queries:
                try:
                    results = list(ddgs.text(query, max_results=3))
                    for r in results:
                        title = r.get("title", "").strip()
                        url   = r.get("href",  "").strip()
                        body  = r.get("body",  "").strip()[:200]
                        if url:
                            all_results.append(
                                f"- [{title}]({url})\n  {body}"
                            )
                except Exception as exc:
                    logger.debug("Search query failed (%r): %s", query, exc)
                    continue

    except Exception as exc:
        logger.warning("DuckDuckGo web search failed: %s", exc)
        return None

    if not all_results:
        return None

    # Deduplicate by URL (keep first occurrence)
    seen_urls: set[str] = set()
    deduped: list[str] = []
    for entry in all_results:
        # Extract URL from markdown link pattern
        url_start = entry.find("(") + 1
        url_end   = entry.find(")")
        url = entry[url_start:url_end] if url_start > 0 and url_end > 0 else ""
        if url and url not in seen_urls:
            seen_urls.add(url)
            deduped.append(entry)

    context = (
        f"VERIFIED WEB SEARCH RESULTS for '{machine_name}':\n"
        + "\n".join(deduped[:5])  # top 5 unique results
        + "\n\nUse the URLs above for Website, Product website, and References "
          "fields. Only use URLs that appear in the search results above."
    )

    logger.info(
        "Web search found %d results for %r", len(deduped), machine_name
    )
    return context
