import { getDB } from '../../../lib/db.js';
import { generateEmbeddings } from '../../../lib/embeddings.js';
import { extractTextFromPDF, validatePDF } from '../../../lib/pdf-extractor.js';
import { chunkText, estimateTokenCount } from '../../../lib/text-chunking.js';
import { withCsrf } from '../../../lib/csrf-middleware.js';

const MAX_FILE_SIZE = 25 * 1024 * 1024; // 25MB
const ALLOWED_TYPES = ['application/pdf', 'text/plain'];
const NULL_BYTE = String.fromCharCode(0);

/**
 * POST /api/upload
 *
 * Document Upload & Indexing Endpoint
 *
 * Flow:
 * 1. Decode uploaded file (base64 JSON payload)
 * 2. Extract text (PDF or plain text)
 * 3. Split into overlapping chunks
 * 4. Generate embeddings for each chunk (Groq)
 * 5. Store document + chunks in Postgres
 */
async function handler(req) {
  try {
    const { fileData, fileName, fileType, sessionId } = await req.json();

    if (!sessionId) {
      return new Response(
        JSON.stringify({ error: 'Session ID is required' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    if (!fileData || !fileName) {
      return new Response(
        JSON.stringify({ error: 'File data and file name are required' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    if (!ALLOWED_TYPES.includes(fileType)) {
      return new Response(
        JSON.stringify({ error: 'Only PDF and TXT files are supported' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    const buffer = Buffer.from(fileData, 'base64');

    if (buffer.length === 0) {
      return new Response(
        JSON.stringify({ error: 'File is empty' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    if (buffer.length > MAX_FILE_SIZE) {
      return new Response(
        JSON.stringify({ error: 'File too large (max 25MB)' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    console.log(`[Upload] Processing "${fileName}" (${buffer.length} bytes) for session ${sessionId}`);

    // Step 1: Extract text
    let text;
    if (fileType === 'application/pdf') {
      const validation = validatePDF({ type: fileType, size: buffer.length });
      if (!validation.valid) {
        return new Response(
          JSON.stringify({ error: validation.error }),
          { status: 400, headers: { 'Content-Type': 'application/json' } }
        );
      }
      const extracted = await extractTextFromPDF(buffer);
      text = extracted.text;
    } else {
      text = buffer.toString('utf-8');
    }

    // Strip null bytes - Postgres text columns reject them
    text = text.split(NULL_BYTE).join('');

    if (!text || text.trim().length === 0) {
      return new Response(
        JSON.stringify({ error: 'No text could be extracted from this file' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    // Step 2: Chunk text
    const chunks = chunkText(text, { chunkSize: 500, overlap: 50 });

    if (chunks.length === 0) {
      return new Response(
        JSON.stringify({ error: 'No content to index after chunking' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    console.log(`[Upload] Split into ${chunks.length} chunks`);

    // Step 3: Generate embeddings (batched)
    const embeddings = await generateEmbeddings(chunks);

    console.log(`[Upload] Generated ${embeddings.length} embeddings`);

    // Step 4: Store document + chunks
    const db = getDB();

    const documentResult = await db`
      INSERT INTO documents (name, file_type, file_size, session_id, chunk_count)
      VALUES (${fileName}, ${fileType}, ${buffer.length}, ${sessionId}, ${chunks.length})
      RETURNING id, name
    `;

    const document = documentResult.rows[0];

    await Promise.all(
      chunks.map((chunk, index) =>
        db`
          INSERT INTO chunks (document_id, chunk_index, content, token_count, embedding)
          VALUES (
            ${document.id},
            ${index},
            ${chunk},
            ${estimateTokenCount(chunk)},
            ${JSON.stringify(embeddings[index])}
          )
        `
      )
    );

    console.log(`[Upload] Document "${fileName}" indexed successfully`);

    return new Response(
      JSON.stringify({
        document: {
          id: document.id,
          name: document.name,
          chunkCount: chunks.length
        }
      }),
      { status: 200, headers: { 'Content-Type': 'application/json' } }
    );

  } catch (error) {
    console.error('[Upload] Error processing upload:', error);
    return new Response(
      JSON.stringify({
        error: 'Failed to process upload',
        details: error.message
      }),
      { status: 500, headers: { 'Content-Type': 'application/json' } }
    );
  }
}

export const POST = withCsrf(handler);
