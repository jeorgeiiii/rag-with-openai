import OpenAI from 'openai';

/**
 * Gemini Embeddings Utility (OpenAI-compatible endpoint)
 *
 * Groq has no embedding models, so embeddings run on Gemini's free tier
 * (chat completions still run on Groq - see app/api/chat/route.js)
 *
 * Generates 768-dimensional embeddings using gemini-embedding-001.
 * The model supports flexible dimensions (128-3072); we request 768 to
 * match the chunks.embedding column, and truncate defensively in case
 * the OpenAI-compat layer doesn't honor the `dimensions` param.
 */

const gemini = new OpenAI({
  apiKey: process.env.GEMINI_API_KEY,
  baseURL: 'https://generativelanguage.googleapis.com/v1beta/openai/'
});

const EMBEDDING_MODEL = 'gemini-embedding-001';
const EMBEDDING_DIMENSIONS = 768;

function toFixedDimensions(vector) {
  return vector.length > EMBEDDING_DIMENSIONS
    ? vector.slice(0, EMBEDDING_DIMENSIONS)
    : vector;
}

/**
 * Generate embedding for a single text chunk
 *
 * @param {string} text - Text to embed
 * @returns {Promise<number[]>} - 768-dimensional embedding vector
 */
export async function generateEmbedding(text) {
  try {
    const response = await gemini.embeddings.create({
      model: EMBEDDING_MODEL,
      input: text,
      dimensions: EMBEDDING_DIMENSIONS,
      encoding_format: 'float'
    });

    return toFixedDimensions(response.data[0].embedding);
  } catch (error) {
    console.error('[Embeddings] Error generating embedding:', error);
    throw new Error('Failed to generate embedding');
  }
}

/**
 * Generate embeddings for multiple text chunks (batch processing)
 *
 * @param {string[]} texts - Array of texts to embed
 * @returns {Promise<number[][]>} - Array of 768-dimensional embedding vectors
 */
export async function generateEmbeddings(texts) {
  try {
    const response = await gemini.embeddings.create({
      model: EMBEDDING_MODEL,
      input: texts,
      dimensions: EMBEDDING_DIMENSIONS,
      encoding_format: 'float'
    });

    return response.data.map(item => toFixedDimensions(item.embedding));
  } catch (error) {
    console.error('[Embeddings] Error generating embeddings:', error);
    throw new Error('Failed to generate embeddings');
  }
}

/**
 * Calculate cosine similarity between two vectors
 * (For testing/debugging - Postgres pgvector handles this in production)
 *
 * @param {number[]} a - First vector
 * @param {number[]} b - Second vector
 * @returns {number} - Similarity score (0-1, higher = more similar)
 */
export function cosineSimilarity(a, b) {
  let dotProduct = 0;
  let magnitudeA = 0;
  let magnitudeB = 0;

  for (let i = 0; i < a.length; i++) {
    dotProduct += a[i] * b[i];
    magnitudeA += a[i] * a[i];
    magnitudeB += b[i] * b[i];
  }

  magnitudeA = Math.sqrt(magnitudeA);
  magnitudeB = Math.sqrt(magnitudeB);

  if (magnitudeA === 0 || magnitudeB === 0) return 0;

  return dotProduct / (magnitudeA * magnitudeB);
}
