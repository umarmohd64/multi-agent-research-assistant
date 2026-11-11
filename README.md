# Lumen Research Studio

Lumen is a multi-agent research assistant that searches the web, reads sources, writes a report, and critiques the result.

**Hosted app:** [lumen-studio-ai.streamlit.app](https://lumen-studio-ai.streamlit.app/)

## Features

- Live search, source extraction, report writing, and critique
- Agent progress, research lounge, and light/dark themes
- Downloadable Word reports and session-only research history

## Run locally

1. Create a Python virtual environment and install dependencies:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

2. Create a `.env` file in the project folder with your API keys:

   ```text
   OPENROUTER_API_KEY=your_openrouter_api_key
   TAVILY_API_KEY=your_tavily_api_key
   ```

3. Start the app:

   ```powershell
   .\.venv\Scripts\python.exe -m streamlit run app.py
   ```

The app uses OpenRouter's `openai/gpt-4o-mini` model and Tavily search. API usage may incur charges. Keep your keys private.

---
Made by Mohd Umar | 2026
