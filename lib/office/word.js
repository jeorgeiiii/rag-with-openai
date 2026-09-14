import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { runComAction } from './com-runner.js';
import { resolveSafePath } from './workdir.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCRIPT_PATH = path.resolve(__dirname, '../../scripts/office/word-com.ps1');

function runWordAction(action, payload) {
  return runComAction(SCRIPT_PATH, 'Word', action, payload);
}

export async function createDocument({ fileName, title }) {
  const filePath = resolveSafePath(fileName, 'docx');
  if (fs.existsSync(filePath)) {
    throw new Error(`"${path.basename(filePath)}" already exists. Choose a different name.`);
  }
  return runWordAction('CreateDocument', { filePath, title: title || '' });
}

export async function addParagraph({ fileName, text, style }) {
  return runWordAction('AddParagraph', {
    filePath: resolveSafePath(fileName, 'docx'),
    text,
    style: style || ''
  });
}

export async function readText({ fileName }) {
  return runWordAction('ReadText', { filePath: resolveSafePath(fileName, 'docx') });
}

export async function replaceText({ fileName, findText, replaceText: replacement }) {
  return runWordAction('ReplaceText', {
    filePath: resolveSafePath(fileName, 'docx'),
    findText,
    replaceText: replacement
  });
}

// OpenAI/Groq-compatible tool (function calling) definitions.
export const WORD_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'create_document',
      description: 'Create a new Word document (.docx) in the chatbot office folder, optionally with a title. Fails if the file already exists.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string', description: 'File name, e.g. "Project Proposal". .docx is added automatically.' },
          title: { type: 'string', description: 'Optional title text added as the first, Title-styled paragraph.' }
        },
        required: ['fileName']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_paragraph',
      description: 'Append a paragraph of text to the end of an existing Word document.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          text: { type: 'string' },
          style: { type: 'string', description: 'Optional Word paragraph style name, e.g. "Heading 1", "Heading 2", "Normal".' }
        },
        required: ['fileName', 'text']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_document_text',
      description: 'Read the full text content of a Word document.',
      parameters: {
        type: 'object',
        properties: { fileName: { type: 'string' } },
        required: ['fileName']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'replace_text',
      description: 'Find and replace all occurrences of a piece of text in a Word document.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          findText: { type: 'string' },
          replaceText: { type: 'string' }
        },
        required: ['fileName', 'findText', 'replaceText']
      }
    }
  }
];

export const WORD_TOOL_IMPL = {
  create_document: createDocument,
  add_paragraph: addParagraph,
  read_document_text: readText,
  replace_text: replaceText
};
