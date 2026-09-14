import OpenAI from 'openai';
import { withCsrf } from '../../../lib/csrf-middleware.js';
import { EXCEL_TOOLS, EXCEL_TOOL_IMPL } from '../../../lib/office/excel.js';
import { WORD_TOOLS, WORD_TOOL_IMPL } from '../../../lib/office/word.js';
import { PPTX_TOOLS, PPTX_TOOL_IMPL } from '../../../lib/office/pptx.js';
import { ACCESS_TOOLS, ACCESS_TOOL_IMPL } from '../../../lib/office/access.js';
import { PUBLISHER_TOOLS, PUBLISHER_TOOL_IMPL } from '../../../lib/office/publisher.js';
import { getWorkdir } from '../../../lib/office/workdir.js';

const groq = new OpenAI({
  apiKey: process.env.GROQ_API_KEY,
  baseURL: 'https://api.groq.com/openai/v1'
});

const MODEL = 'openai/gpt-oss-120b';
const MAX_TOOL_ROUNDS = 6;

const ALL_TOOLS = [...EXCEL_TOOLS, ...WORD_TOOLS, ...PPTX_TOOLS, ...ACCESS_TOOLS, ...PUBLISHER_TOOLS];
const ALL_TOOL_IMPL = { ...EXCEL_TOOL_IMPL, ...WORD_TOOL_IMPL, ...PPTX_TOOL_IMPL, ...ACCESS_TOOL_IMPL, ...PUBLISHER_TOOL_IMPL };

/**
 * POST /api/office-chat
 *
 * Natural-language -> Excel automation, via tool calling.
 * Windows + Excel installed locally only; not available on the Vercel deployment.
 */
async function handler(req) {
  try {
    if (process.platform !== 'win32') {
      return new Response(
        JSON.stringify({
          error: 'The Office Assistant requires Windows with Microsoft Office installed. It only runs when this app is started locally (npm run dev / npm start) on that machine - it is not available on this deployment.'
        }),
        { status: 501, headers: { 'Content-Type': 'application/json' } }
      );
    }

    const { message, history } = await req.json();

    if (!message || !message.trim()) {
      return new Response(JSON.stringify({ error: 'message is required' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const systemPrompt = `You are an Office assistant that edits real Excel, Word, PowerPoint, Access, and Publisher files on the user's computer using tools.
All files live in a single folder: ${getWorkdir()}. Refer to files by name only, not full paths.
Pick the right tool set based on what the user wants: spreadsheets/workbooks -> Excel tools, documents -> Word tools, slides/decks -> PowerPoint tools, databases/tables -> Access tools, flyers/publications -> Publisher tools.
A file must be created (create_workbook / create_document / create_presentation / create_database / create_publication) before it can be edited or read.
Always confirm what you did in plain language after using tools (e.g. which file and what changed).
If a request is ambiguous (e.g. no file name given), ask a brief clarifying question instead of guessing.
Keep replies short.`;

    const messages = [
      { role: 'system', content: systemPrompt },
      ...(Array.isArray(history) ? history : []),
      { role: 'user', content: message }
    ];

    const actionsPerformed = [];

    for (let round = 0; round < MAX_TOOL_ROUNDS; round++) {
      const completion = await groq.chat.completions.create({
        model: MODEL,
        messages,
        tools: ALL_TOOLS,
        temperature: 0.2,
        max_tokens: 400
      });

      const choice = completion.choices[0];
      const assistantMessage = choice.message;
      messages.push(assistantMessage);

      const toolCalls = assistantMessage.tool_calls || [];
      if (toolCalls.length === 0) {
        return new Response(
          JSON.stringify({
            response: assistantMessage.content || '',
            actions: actionsPerformed
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } }
        );
      }

      for (const call of toolCalls) {
        const name = call.function?.name;
        let args = {};
        try {
          args = JSON.parse(call.function?.arguments || '{}');
        } catch {
          // leave args empty; the tool impl will surface a clear error
        }

        const impl = ALL_TOOL_IMPL[name];
        let resultPayload;
        let errored = false;

        if (!impl) {
          errored = true;
          resultPayload = { ok: false, error: `Unknown tool: ${name}` };
        } else {
          try {
            resultPayload = await impl(args);
          } catch (err) {
            errored = true;
            resultPayload = { ok: false, error: err.message };
          }
        }

        actionsPerformed.push({ tool: name, args, ok: !errored, result: resultPayload });

        messages.push({
          role: 'tool',
          tool_call_id: call.id,
          content: JSON.stringify(resultPayload)
        });
      }
    }

    return new Response(
      JSON.stringify({
        response: "I've made several changes but stopped after a few steps — ask me to continue if needed.",
        actions: actionsPerformed
      }),
      { status: 200, headers: { 'Content-Type': 'application/json' } }
    );
  } catch (error) {
    console.error('[OfficeChat] Error:', error);
    return new Response(
      JSON.stringify({ error: 'Failed to process request', details: error.message }),
      { status: 500, headers: { 'Content-Type': 'application/json' } }
    );
  }
}

export const POST = withCsrf(handler);
