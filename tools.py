from langchain.tools import tool
import requests
from bs4 import BeautifulSoup
from tavily import TavilyClient
from rich import print
import os
from typing import Literal
from dotenv import load_dotenv
load_dotenv()

tavily = TavilyClient(api_key=os.getenv("TAVILY_API_KEY"))

@tool
def web_search(
    query: str,
    topic: Literal["general", "news"] = "general",
    time_range: Literal["day", "week", "month", "year"] | None = None,
) -> str:
    """Search the web. Use topic='news' and time_range='day' for latest news.
    Broaden the time range only when needed and disclose it. For historical or
    general research, use topic='general' and omit time_range when appropriate.
    """
    results = tavily.search(
        query=query,
        topic=topic,
        time_range=time_range,
        max_results=6,
        include_published_date=True,
        filter_by_published_date=time_range is not None,
    )

    out = []

    for r in results['results']:
        out.append(
            f"Title: {r['title']},\nURL: {r['url']},\n"
            f"Published or updated: {r.get('published_date') or 'Unknown'},\n"
            f"Snippet: {r['content'][:1200]}\n"
        )

    if not out:
        return "No results found for this query and time range."
    return "\n----\n".join(out)


@tool
def scrape_url(url: str) -> str:
    """Scrape the content of a URL and return the clean text for deeper reading, information and analysis."""
    try:
        response = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        return soup.get_text(separator=" ", strip=True)[:4000]
    except requests.RequestException as e:
        return f"Error scraping URL: {str(e)}"




