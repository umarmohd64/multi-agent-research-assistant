from agents import build_search_agent, build_scraper_agent, writer_chain, critic_chain
from datetime import datetime
from rich import print

def run_research_pipeline(topic: str) -> dict:

    state= {}
    current_date = datetime.now().astimezone().strftime("%Y-%m-%d")

    # search agnet working
    print("\n"+" ="*50)
    print("Running search agent...")
    print("="*50)

    search_agent = build_search_agent()
    search_results = search_agent.invoke({
        "messages": [("user", (
            f"Today's local date is {current_date}. Research this topic: {topic}\n"
            "Use the web search tool; do not rely on memory. For current news or "
            "latest developments, search with topic='news' and time_range='day' "
            "first. If there are insufficient relevant results, search the week "
            "and then the month as needed, and disclose the range used. For other "
            "topics, use an appropriate search category and date range, including "
            "historical sources when the topic asks for them. Prefer reliable, "
            "independent sources and corroborate important claims. Report each "
            "source's title, URL, publication or update date, and relevant facts. "
            "Separate current developments from older background, and say when "
            "dates or sufficient recent evidence are unavailable. Never invent sources."
        ))]
    })

    state["searched_content"] = search_results['messages'][-1].content

    print("\nSearch Results:", state["searched_content"])



    # scraper agent working
    print("\n"+" ="*50)
    print("Running scraper agent for extracting resources...")
    print("="*50)

    scraper_agent = build_scraper_agent()
    scrape_results = scraper_agent.invoke({
        "messages": [("user", 
                      f"Today's local date is {current_date}. For the topic {topic}, "
                      "use the scraping tool on the most relevant, recent, reliable "
                      "source URLs below. Extract source-specific facts and preserve "
                      "each URL and any publication, update, and event dates. "
                      "Identify older background and conflicting accounts; do not "
                      "invent facts or treat failed scrapes as sources.\n"
                      f"Search results:\n{state['searched_content']}"
                    )]
    })

    state["scraped_content"] = scrape_results['messages'][-1].content

    print("\nScrape Results:", state["scraped_content"])


    # Writer Chain
    print("\n"+" ="*50)
    print("Running writer chain for generating content...")
    print("="*50)

    combined_content = (
        f"SEARCH RESULTS: \n {state['searched_content']}\n\n"
        f"SCRAPED CONTENT: \n {state['scraped_content']}\n"
        )

    state["report"] = writer_chain.invoke({
        "topic": topic,
        "research": combined_content,
        "current_date": current_date,
    })

    print("\nFinal Report Generated:", state["report"])


    # Critic Chain
    print("\n"+" ="*50)
    print("Running critic chain for evaluating the report...")
    print("="*50)

    state["feedback"] = critic_chain.invoke({
        "report": state["report"],
        "research": combined_content,
        "current_date": current_date,
    })

    print("\nCritique of the Report:", state["feedback"])


    return state


if __name__ == "__main__":
    topic = input("Enter the research topic: ")
    run_research_pipeline(topic)





