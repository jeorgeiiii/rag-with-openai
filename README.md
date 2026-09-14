# RAG Chatbot + Microsoft Office Assistant

A Retrieval-Augmented Generation chatbot — upload PDFs or text files, and ask questions that get answered from *those documents*, with source citations, instead of the model's training data — **plus a second assistant that drives real Microsoft Office applications from natural language**: create and edit Excel workbooks, Word documents, PowerPoint decks, Access databases, and Publisher publications, all by chatting.

Built to run entirely on free-tier services: **Neon** (Postgres + pgvector) for storage, **Groq** for chat generation, and **Gemini** for embeddings, since Groq doesn't offer an embeddings API of its own. The Office Assistant uses **Windows COM automation** — no Microsoft Graph API, no Azure app registration, no paid API tier.

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

## Office Assistant — control Excel, Word, PowerPoint, Access & Publisher from chat

A second, independent assistant at **`/office`** turns plain-English requests into real edits on real Office files, using **LLM tool calling** (Groq) to pick the right action, backed by **Windows COM automation** rather than a cloud API:

```
"Create a workbook called Budget, write headers Name and Total,       Groq picks a tool         PowerShell drives Excel/Word/
 then sum column B into C1"                                    →      (function calling)   →    PowerPoint/Access/Publisher
                                                                                                   via COM, headless
```

Each Office app has its own tool set the model chooses from based on intent:

| App | Actions | Example |
|---|---|---|
| **Excel** | create workbook, add sheet, write range, read range, apply formula | *"Create Sales.xlsx with a Q1 sheet, write these rows, sum column B"* |
| **Word** | create document, append paragraph, read text, find & replace | *"Create a doc titled Meeting Notes, add a paragraph about the roadmap"* |
| **PowerPoint** | create presentation, add slide with bullets, list/read slides | *"Create a deck titled Our Product, add a slide with 3 bullet points"* |
| **Access** | create database, create table, insert rows, run SELECT queries | *"Create a People table with Name and Age, insert two rows, show everyone over 25"* |
| **Publisher** | create publication, add text box, read text | *"Create a flyer with a headline, add a text box with the offer"* |

**Every file is sandboxed** to one folder (`~/Documents/ChatbotOffice` by default, override with `OFFICE_WORKDIR`) — file names are resolved and validated server-side so a request can never read or write outside it, no matter what the model is asked to do.

**Not (yet) supported, and why:**

| App | Status |
|---|---|
| **Outlook** | Planned, drafts-only by design (no tool would ever be able to send an email — creating/saving a draft only). Not wired up yet — requires a configured mail profile, which is a one-time interactive step only the machine's owner can complete. |
| **OneNote** | Notebook/section/page creation works over COM; setting page title/body text does not (`UpdatePageContent` fails with `HRESULT 0x80042030` across every documented approach tried) — looks like an environment-side restriction, not something fixable client-side. |
| **Teams, SharePoint, OneDrive, Forms** | No desktop COM automation model exists for these — they're cloud services. Supporting them would mean switching to the Microsoft Graph API (Azure app registration + OAuth), a different, non-free-tier architecture from the rest of this project. |
| **Visio, Project** | Not installed / separate license — not tested. |

**Requirements**: Windows + the relevant Office app installed and licensed on the machine running the server. This only works when you run the app locally (`npm run dev` / `npm start`) — **it does not and cannot run on the Vercel deployment**, which is Linux with no Office installed. The code detects this itself: every Office action checks `process.platform === 'win32'` first and fails safe with a clear message instead of erroring obscurely.

### Why COM automation instead of Microsoft Graph API

Graph API is the "proper" cloud way to do this, but it needs an Azure app registration, OAuth consent, and files living in OneDrive/SharePoint. COM automation needs none of that — it drives the Office apps already installed on the machine directly, for free, with no cloud round-trip. The trade-off: it only works on a Windows box with Office installed and running the server process, which is an intentional fit for a personal/local-first tool rather than a multi-tenant SaaS.

### Engineering notes

Getting reliable COM automation working surfaced several real, non-obvious bugs along the way — each isolated with a minimal repro before being fixed:

- **Word's `Document.SaveAs`/`SaveAs2` hangs indefinitely** on first save of a new document, regardless of content, format, or destination — reproduced with a bare `Documents.Add()` + `SaveAs2()`, confirmed `Documents.Open()` + `.Save()` on an existing file works fine. Worked around by building new `.docx` files directly as minimal OOXML packages (`System.IO.Packaging`) and only ever using the reliable Open+Save path afterward.
- **PowerShell's COM property-chaining silently fails** on patterns like `$rs.Fields.Item(name).Value = x` — fixed by binding the intermediate COM object to a variable first, and in Access's case by falling back to late-bound reflection (`InvokeMember`) to dodge a spurious `InvalidCastException` from early-bound type resolution.
- **A false-timeout bug in the tool runner**: some apps (Publisher) write their result file successfully but are slow to fully quit/release their process. The original timeout logic killed the process and reported failure even though the work had already completed — fixed by checking for a completed result file before treating a timeout as a real failure.

## Tech stack

| Layer | Choice |
|---|---|
| Framework | Next.js 16 (App Router), React 19 |
| Database | Neon Postgres + `pgvector` (HNSW index, cosine distance) |
| Chat model | Groq — `openai/gpt-oss-120b` |
| Embeddings | Google Gemini — `gemini-embedding-001`, 768 dimensions |
| PDF parsing | `pdf-parse` v2 (pdfjs-dist under the hood) |
| Office automation | Windows COM via `powershell.exe`, driven by Groq tool calling |
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
| `OFFICE_WORKDIR` | no | Folder the Office Assistant may read/write `.xlsx`/`.docx`/`.pptx`/`.accdb`/`.pub` files in. Defaults to `~/Documents/ChatbotOffice`. Windows + Office only — see [Office Assistant](#office-assistant--control-excel-word-powerpoint-access--publisher-from-chat) |

## Project structure

```
rag-chatbot/
├── app/
│   ├── api/
│   │   ├── chat/            # POST — embed query, vector search, generate answer
│   │   ├── upload/          # POST — extract, chunk, embed, store a document
│   │   ├── documents/       # GET  — list uploaded documents
│   │   ├── clear-session/   # POST — wipe a session's documents/chunks (cascades)
│   │   ├── csrf-token/      # GET  — issue a CSRF token + cookie
│   │   └── office-chat/     # POST — tool-calling loop over the Office COM tools
│   ├── office/               # /office — Office Assistant page
│   ├── layout.js
│   └── page.js
├── components/
│   ├── ChatInterface.js     # RAG chat UI, upload flow, document list
│   ├── OfficeInterface.js   # Office Assistant chat UI
│   ├── ConfirmModal.js
│   ├── WelcomeModal.js
│   └── icons/                # inline SVG icon components
├── lib/
│   ├── db.js                 # Neon connection (DATABASE_URL → POSTGRES_URL)
│   ├── embeddings.js         # Gemini embeddings client
│   ├── pdf-extractor.js      # pdf-parse wrapper
│   ├── text-chunking.js      # sentence/paragraph-aware chunker
│   ├── csrf.js                # token generation/verification
│   ├── csrf-middleware.js     # withCsrf() route wrapper
│   └── office/
│       ├── workdir.js         # sandboxed path resolution for Office files
│       ├── com-runner.js      # spawns powershell.exe, parses JSON result
│       ├── excel.js           # Excel tool defs + implementations
│       ├── word.js            # Word tool defs + implementations
│       ├── pptx.js            # PowerPoint tool defs + implementations
│       ├── access.js          # Access tool defs + implementations
│       └── publisher.js       # Publisher tool defs + implementations
├── scripts/
│   ├── init-db.js             # creates tables, HNSW index, document_stats view
│   ├── add-session-isolation.js  # adds session_id + 24h cleanup function
│   └── office/
│       ├── excel-com.ps1      # Excel COM automation (one action per invocation)
│       ├── word-com.ps1       # Word COM automation
│       ├── powerpoint-com.ps1 # PowerPoint COM automation
│       ├── access-com.ps1     # Access/DAO automation
│       └── publisher-com.ps1  # Publisher COM automation
└── ...
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

**`POST /api/office-chat`** — Windows + Office only, see [Office Assistant](#office-assistant--control-excel-word-powerpoint-access--publisher-from-chat)
```json
{ "message": "create a workbook called Budget with a Sales sheet", "history": [] }
→ { "response": "...", "actions": [{ "tool": "create_workbook", "args": {...}, "ok": true, "result": {...} }] }
```

## Database schema

- **`documents`** — `id`, `name`, `file_type`, `file_size`, `session_id`, `chunk_count`, `metadata`, timestamps
- **`chunks`** — `id`, `document_id` (FK, `ON DELETE CASCADE`), `chunk_index`, `content`, `token_count`, `embedding vector(768)`
- **`query_history`** — `query`, `response`, `retrieved_chunks`, timing columns (retrieval/generation/total ms)
- Indexes: `idx_chunks_document_id`, `idx_chunks_embedding` (HNSW, `vector_cosine_ops`, `m=16, ef_construction=64`), `idx_documents_session_id`

## Security

**RAG chatbot:**
- CSRF: double-submit cookie pattern, constant-time comparison
- Session isolation: every query/delete is scoped to `session_id`; no cross-session data access
- Parameterized SQL throughout (tagged-template queries via `@vercel/postgres`, no string interpolation)
- Null bytes stripped from extracted text before insert (Postgres `text` rejects them)
- Generic error responses to clients; details logged server-side only

**Office Assistant — who can actually use it:**

The Office Assistant only ever acts on **the single machine running the Node server** — there's no mechanism by which a visitor's own computer or Office installation is touched.

- **On Vercel**: every Office action checks `process.platform === 'win32'` first and refuses to run on Linux, so the public deployment is safe by construction — anyone can use the RAG chatbot there, but the Office Assistant simply returns an error for everyone.
- **Run locally and not exposed**: private to whoever is at that machine; nobody else can reach it.
- **⚠️ If you tunnel/expose your local server** (ngrok, port forwarding, etc.): `/api/office-chat` currently has **no authentication**, only CSRF protection (which stops forged cross-site requests, not a real visitor who loaded the page directly). Anyone with the link could create, read, or edit Excel/Word/PowerPoint/Access/Publisher files on your machine. File access is sandboxed to one folder (`OFFICE_WORKDIR`, path-validated server-side against traversal) so it can't reach the rest of your filesystem — but within that folder, it's unauthenticated. Don't expose it publicly without adding a login in front of it first.

Other guardrails already in place:
- Every file path is resolved and validated against the sandbox folder before any COM action runs — `../` and absolute paths are rejected.
- Access's `query_table` tool only accepts `SELECT` statements (regex-enforced); data changes go through `insert_rows`, which uses parameterized recordset writes, not string-built SQL.

## Known limitation: upload size on Vercel

Vercel serverless functions cap request bodies at **4.5MB**, regardless of the app's own 25MB check. Base64-encoding a file inflates its size ~33%, so PDFs beyond roughly 3MB will fail to upload once deployed, even though they work fine locally. Set `NEXT_PUBLIC_UPLOAD_URL` to route uploads through an external service to work around this, or wire up `@vercel/blob` (already a dependency, not yet used) for direct client-side uploads.

## Deployment

```bash
npm i -g vercel
vercel --prod
```

Set `DATABASE_URL`, `GROQ_API_KEY`, `GEMINI_API_KEY`, and `SESSION_SECRET` in the Vercel project's environment variables. If deploying against a fresh database, run `npm run db:init` and `node scripts/add-session-isolation.js` against it first.

Only the RAG chatbot deploys to Vercel this way. The Office Assistant (`/office`) is local-only by design — see [Office Assistant](#office-assistant--control-excel-word-powerpoint-access--publisher-from-chat) and [Security](#security).

## License

MIT

## Author

**Prince Mehra** — [Portfolio](https://portfolio-ruddy-seven-7slackrg3a.vercel.app/) · [GitHub](https://github.com/jeorgeiiii) · [LinkedIn](https://www.linkedin.com/in/prince-mehra-b3322935a/)
