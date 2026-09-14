import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { runComAction } from './com-runner.js';
import { resolveSafePath } from './workdir.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCRIPT_PATH = path.resolve(__dirname, '../../scripts/office/publisher-com.ps1');

function runPublisherAction(action, payload) {
  return runComAction(SCRIPT_PATH, 'Publisher', action, payload);
}

export async function createPublication({ fileName, text }) {
  const filePath = resolveSafePath(fileName, 'pub');
  if (fs.existsSync(filePath)) {
    throw new Error(`"${path.basename(filePath)}" already exists. Choose a different name.`);
  }
  return runPublisherAction('CreatePublication', { filePath, text: text || '' });
}

export async function addTextBox({ fileName, text, pageIndex, left, top, width, height }) {
  return runPublisherAction('AddTextBox', {
    filePath: resolveSafePath(fileName, 'pub'),
    text,
    pageIndex,
    left,
    top,
    width,
    height
  });
}

export async function readText({ fileName }) {
  return runPublisherAction('ReadText', { filePath: resolveSafePath(fileName, 'pub') });
}

// OpenAI/Groq-compatible tool (function calling) definitions.
export const PUBLISHER_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'create_publication',
      description: 'Create a new Publisher publication (.pub) in the chatbot office folder, optionally with an initial text box. Fails if the file already exists.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string', description: 'File name, e.g. "Flyer". .pub is added automatically.' },
          text: { type: 'string', description: 'Optional text for an initial text box on page 1.' }
        },
        required: ['fileName']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_text_box',
      description: 'Add a text box to a page in an existing publication.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          text: { type: 'string' },
          pageIndex: { type: 'integer', description: '1-based page number, defaults to 1.' },
          left: { type: 'number', description: 'Position in points from the left, defaults to 72.' },
          top: { type: 'number', description: 'Position in points from the top, defaults to 72.' },
          width: { type: 'number', description: 'Width in points, defaults to 300.' },
          height: { type: 'number', description: 'Height in points, defaults to 100.' }
        },
        required: ['fileName', 'text']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_publication_text',
      description: 'Read all text box content across every page of a publication.',
      parameters: {
        type: 'object',
        properties: { fileName: { type: 'string' } },
        required: ['fileName']
      }
    }
  }
];

export const PUBLISHER_TOOL_IMPL = {
  create_publication: createPublication,
  add_text_box: addTextBox,
  read_publication_text: readText
};
