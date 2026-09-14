import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { runComAction } from './com-runner.js';
import { resolveSafePath } from './workdir.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCRIPT_PATH = path.resolve(__dirname, '../../scripts/office/powerpoint-com.ps1');

function runPptAction(action, payload) {
  return runComAction(SCRIPT_PATH, 'PowerPoint', action, payload);
}

export async function createPresentation({ fileName, title, subtitle }) {
  const filePath = resolveSafePath(fileName, 'pptx');
  if (fs.existsSync(filePath)) {
    throw new Error(`"${path.basename(filePath)}" already exists. Choose a different name.`);
  }
  return runPptAction('CreatePresentation', { filePath, title: title || '', subtitle: subtitle || '' });
}

export async function addSlide({ fileName, title, bullets }) {
  return runPptAction('AddSlide', {
    filePath: resolveSafePath(fileName, 'pptx'),
    title: title || '',
    bullets: bullets || []
  });
}

export async function listSlides({ fileName }) {
  return runPptAction('ListSlides', { filePath: resolveSafePath(fileName, 'pptx') });
}

export async function readSlide({ fileName, slideIndex }) {
  return runPptAction('ReadSlide', { filePath: resolveSafePath(fileName, 'pptx'), slideIndex });
}

// OpenAI/Groq-compatible tool (function calling) definitions.
export const PPTX_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'create_presentation',
      description: 'Create a new PowerPoint presentation (.pptx) in the chatbot office folder with a title slide. Fails if the file already exists.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string', description: 'File name, e.g. "Q1 Review". .pptx is added automatically.' },
          title: { type: 'string', description: 'Title slide main text.' },
          subtitle: { type: 'string', description: 'Optional title slide subtitle text.' }
        },
        required: ['fileName', 'title']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_slide',
      description: 'Append a new slide with a title and bullet-point body to an existing presentation.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          title: { type: 'string' },
          bullets: {
            type: 'array',
            items: { type: 'string' },
            description: 'Bullet point lines for the slide body, e.g. ["Revenue up 12%", "Churn down 3%"]'
          }
        },
        required: ['fileName', 'title']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'list_slides',
      description: 'List slide titles and count in a presentation.',
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
      name: 'read_slide',
      description: 'Read all text content on a specific slide (1-indexed).',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          slideIndex: { type: 'integer', description: '1-based slide number' }
        },
        required: ['fileName', 'slideIndex']
      }
    }
  }
];

export const PPTX_TOOL_IMPL = {
  create_presentation: createPresentation,
  add_slide: addSlide,
  list_slides: listSlides,
  read_slide: readSlide
};
