import { getDB } from '../../../lib/db.js';

/**
 * GET /api/documents
 *
 * List all uploaded documents with metadata
 */
export async function GET(req) {
  try {
    const { searchParams } = new URL(req.url);
    const sessionId = searchParams.get('sessionId');

    if (!sessionId) {
      return new Response(
        JSON.stringify({ error: 'sessionId is required' }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      );
    }

    const db = getDB();

    const documentsResult = await db`
      SELECT
        id,
        name,
        file_type,
        file_size,
        chunk_count,
        upload_date,
        metadata
      FROM documents
      WHERE session_id = ${sessionId}
      ORDER BY upload_date DESC
    `;

    return new Response(
      JSON.stringify({ documents: documentsResult.rows }),
      { status: 200, headers: { 'Content-Type': 'application/json' } }
    );
  } catch (error) {
    console.error('[Documents] Error fetching documents:', error);
    return new Response(
      JSON.stringify({ error: 'Failed to fetch documents' }),
      { status: 500, headers: { 'Content-Type': 'application/json' } }
    );
  }
}
