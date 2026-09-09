# RAG Chatbot

A Retrieval-Augmented Generation chatbot: upload PDFs or text files, and ask questions that get answered from *those documents* — with source citations — instead of the model's training data.

Built to run entirely on free-tier services: **Neon** (Postgres + pgvector) for storage, **Groq** for chat generation, and **Gemini** for embeddings, since Groq doesn't offer an embeddings API of its own.

## How it works

**Upload**

```
PDF/TXT → extract text → chunk (500 tokens, 50-token overlap) → embed (Gemini, 768-dim) → store in Postgres
```

**Query**

```
question → embed (Gemini) → pgvector cosine search (top 3 chunks, session-scoped) → Groq (openai/gpt-oss-120b) → answer + citations
```

Every document is tagged with a `session_id` generated client-side, so concurrent visitors never see each other's uploads or chat history — there's no login. Sessions and their documents are cleaned up after 24 hours (`scripts/add-session-isolation.js` installs the cleanup function).

## Tech stack

| Layer | Choice |
|---|---|
| Framework | Next.js 16 (App Router), React 19 |
| Database | Neon Postgres + `pgvector` (HNSW index, cosine distance) |
| Chat model | Groq — `openai/gpt-oss-120b` |
| Embeddings | Google Gemini — `gemini-embedding-001`, 768 dimensions |
| PDF parsing | `pdf-parse` v2 (pdfjs-dist under the hood) |
| Styling | Tailwind CSS |
| CSRF | Double-submit cookie, `HttpOnly` + `SameSite=Strict` |

## Getting started

### Prerequisites

- Node.js 18+
- A [Neon](https://neon.tech) Postgres database (pgvector is supported out of the box)
- A [Groq](https://console.groq.com/keys) API key
- A [Gemini](https://aistudio.google.com/apikey) API key

### Setup

```bash
git clone git@github.com:jeorgeiiii/rag-with-openai.git
cd rag-chatbot
npm install

cp .env.local.example .env.local
# fill in DATABASE_URL, GROQ_API_KEY, GEMINI_API_KEY, SESSION_SECRET
```

Initialize the schema, then apply the session-isolation migration:

```bash
npm run db:init
node scripts/add-session-isolation.js
```

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

### Environment variables

| Variable | Required | Notes |
|---|---|---|
| `DATABASE_URL` | yes | Neon connection string; `lib/db.js` maps it to `POSTGRES_URL` for `@vercel/postgres` at runtime |
| `GROQ_API_KEY` | yes | Chat completions |
| `GEMINI_API_KEY` | yes | Embeddings only, via Gemini's OpenAI-compatible endpoint |
| `SESSION_SECRET` | recommended | Random 32+ byte string (`openssl rand -base64 32`) |
| `ENFORCE_CSRF` | no | Set to `false` to bypass CSRF checks locally; leave unset in production |
| `NEXT_PUBLIC_UPLOAD_URL` | no | Point uploads at an external service instead of `/api/upload` (see [Known limitation](#known-limitation-upload-size-on-vercel) below) |

## Project structure

```
rag-chatbot/
├── app/
│   ├── api/
│   │   ├── chat/            # POST — embed query, vector search, generate answer
│   │   ├── upload/          # POST — extract, chunk, embed, store a document
│   │   ├── documents/       # GET  — list uploaded documents
│   │   ├── clear-session/   # POST — wipe a session's documents/chunks (cascades)
│   │   └── csrf-token/      # GET  — issue a CSRF token + cookie
│   ├── layout.js
│   └── page.js
├── components/
│   ├── ChatInterface.js     # chat UI, upload flow, document list
│   ├── ConfirmModal.js
│   ├── WelcomeModal.js
│   └── icons/                # inline SVG icon components
├── lib/
│   ├── db.js                 # Neon connection (DATABASE_URL → POSTGRES_URL)
│   ├── embeddings.js         # Gemini embeddings client
│   ├── pdf-extractor.js      # pdf-parse wrapper
│   ├── text-chunking.js      # sentence/paragraph-aware chunker
│   ├── csrf.js                # token generation/verification
│   └── csrf-middleware.js     # withCsrf() route wrapper
└── scripts/
    ├── init-db.js             # creates tables, HNSW index, document_stats view
    └── add-session-isolation.js  # adds session_id + 24h cleanup function
```

## API

All state-changing endpoints (`POST`) require an `X-CSRF-Token` header matching the `csrf_token` cookie from `GET /api/csrf-token`.

**`POST /api/upload`**
```json
{ "fileData": "<base64>", "fileName": "doc.pdf", "fileType": "application/pdf", "sessionId": "..." }
→ { "document": { "id": 1, "name": "doc.pdf", "chunkCount": 12 } }
```

**`POST /api/chat`**
```json
{ "query": "what does the doc say about X?", "sessionId": "..." }
→ { "response": "...", "sources": [{ "documentName", "similarity", "preview" }], "metadata": { "retrievalTime", "generationTime" } }
```

**`GET /api/documents`** → `{ "documents": [...] }`

**`POST /api/clear-session`**
```json
{ "sessionId": "..." }
→ { "success": true, "deletedDocuments": 3 }
```

## Database schema

- **`documents`** — `id`, `name`, `file_type`, `file_size`, `session_id`, `chunk_count`, `metadata`, timestamps
- **`chunks`** — `id`, `document_id` (FK, `ON DELETE CASCADE`), `chunk_index`, `content`, `token_count`, `embedding vector(768)`
- **`query_history`** — `query`, `response`, `retrieved_chunks`, timing columns (retrieval/generation/total ms)
- Indexes: `idx_chunks_document_id`, `idx_chunks_embedding` (HNSW, `vector_cosine_ops`, `m=16, ef_construction=64`), `idx_documents_session_id`

## Security

- CSRF: double-submit cookie pattern, constant-time comparison
- Session isolation: every query/delete is scoped to `session_id`; no cross-session data access
- Parameterized SQL throughout (tagged-template queries via `@vercel/postgres`, no string interpolation)
- Null bytes stripped from extracted text before insert (Postgres `text` rejects them)
- Generic error responses to clients; details logged server-side only

## Known limitation: upload size on Vercel

Vercel serverless functions cap request bodies at **4.5MB**, regardless of the app's own 25MB check. Base64-encoding a file inflates its size ~33%, so PDFs beyond roughly 3MB will fail to upload once deployed, even though they work fine locally. Set `NEXT_PUBLIC_UPLOAD_URL` to route uploads through an external service to work around this, or wire up `@vercel/blob` (already a dependency, not yet used) for direct client-side uploads.

## Deployment

```bash
npm i -g vercel
vercel --prod
```

Set `DATABASE_URL`, `GROQ_API_KEY`, `GEMINI_API_KEY`, and `SESSION_SECRET` in the Vercel project's environment variables. If deploying against a fresh database, run `npm run db:init` and `node scripts/add-session-isolation.js` against it first.

## License

MIT

## Author

**Prince Mehra** — [Portfolio](https://portfolio-ruddy-seven-7slackrg3a.vercel.app/) · [GitHub](https://github.com/jeorgeiiii) · [LinkedIn](https://www.linkedin.com/in/prince-mehra-b3322935a/)
