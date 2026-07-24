---
title: AfriCareer AI
emoji: 🌍
colorFrom: green
colorTo: yellow
sdk: docker
app_port: 7860
pinned: false
short_description: AI-powered career guidance for African youth (RAG-grounded)
---

# AfriCareer AI

Multilingual, RAG-grounded AI career guidance for African youth — CV/cover-letter
generation, resume analysis, learning resources, and a career assistant grounded in
AfDB / UNICEF / ILO frameworks.

Built with Streamlit + OpenAI (`gpt-4o-mini` + embeddings) + Pinecone (serverless vector DB).

## Deploy on Hugging Face Spaces (Docker SDK)

1. Create a new Space → **SDK: Docker** → **blank**.
2. Push this folder (the `.gitignore` keeps `.env`, analytics, and the `.zipx` out).
3. In **Settings → Variables and secrets**, add:
   - `OPENAI_API_KEY`
   - `PINECONE_API_KEY`
4. The Space builds from the `Dockerfile` and serves on port `7860`.

## Run locally

```bash
cp .env.example .env      # fill in your keys
pip install -r requirements.txt
streamlit run app.py
```

## Notes / limitations

- **Analytics** (`africareer_analytics.json`) is written to the container's ephemeral
  filesystem and resets on rebuild/restart. Move to a managed DB before the pilot.
- **API cost** is unbounded on a public deployment — add rate limiting / usage caps.
- Admin credentials are currently hard-coded in `app.py` — move to secrets before pilot.
