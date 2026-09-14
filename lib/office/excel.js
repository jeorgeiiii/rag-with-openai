import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { runComAction } from './com-runner.js';
import { resolveSafePath } from './workdir.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCRIPT_PATH = path.resolve(__dirname, '../../scripts/office/excel-com.ps1');

function runExcelAction(action, payload) {
  return runComAction(SCRIPT_PATH, 'Excel', action, payload);
}

export async function createWorkbook({ fileName, sheets }) {
  const filePath = resolveSafePath(fileName, 'xlsx');
  if (fs.existsSync(filePath)) {
    throw new Error(`"${path.basename(filePath)}" already exists. Choose a different name.`);
  }
  return runExcelAction('CreateWorkbook', { filePath, sheets: sheets || [] });
}

export async function listSheets({ fileName }) {
  return runExcelAction('ListSheets', { filePath: resolveSafePath(fileName, 'xlsx') });
}

export async function addSheet({ fileName, sheetName }) {
  return runExcelAction('AddSheet', { filePath: resolveSafePath(fileName, 'xlsx'), sheetName });
}

export async function writeData({ fileName, sheetName, startCell, rows }) {
  const { row, col } = cellToRowCol(startCell || 'A1');
  return runExcelAction('WriteData', {
    filePath: resolveSafePath(fileName, 'xlsx'),
    sheetName,
    startRow: row,
    startCol: col,
    rows
  });
}

export async function readData({ fileName, sheetName, range }) {
  return runExcelAction('ReadData', { filePath: resolveSafePath(fileName, 'xlsx'), sheetName, range });
}

export async function applyFormula({ fileName, sheetName, cell, formula }) {
  return runExcelAction('ApplyFormula', {
    filePath: resolveSafePath(fileName, 'xlsx'),
    sheetName,
    cell,
    formula
  });
}

function cellToRowCol(cell) {
  const match = /^([A-Za-z]+)(\d+)$/.exec(cell.trim());
  if (!match) throw new Error(`Invalid cell reference: ${cell}`);
  const [, colLetters, rowStr] = match;
  let col = 0;
  for (const ch of colLetters.toUpperCase()) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return { row: parseInt(rowStr, 10), col };
}

// OpenAI/Groq-compatible tool (function calling) definitions.
export const EXCEL_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'create_workbook',
      description: 'Create a new Excel workbook (.xlsx) in the chatbot office folder. Fails if the file already exists.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string', description: 'File name, e.g. "Sales Report". .xlsx is added automatically.' },
          sheets: {
            type: 'array',
            items: { type: 'string' },
            description: 'Optional list of sheet names to create, e.g. ["Q1", "Q2"]. Defaults to a single "Sheet1".'
          }
        },
        required: ['fileName']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'list_sheets',
      description: 'List the sheet names in an existing workbook.',
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
      name: 'add_sheet',
      description: 'Add a new sheet to an existing workbook.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          sheetName: { type: 'string' }
        },
        required: ['fileName', 'sheetName']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'write_data',
      description: 'Write a 2D grid of values into a sheet, starting at a given cell (e.g. "A1"). Use for tables, headers, and rows of data.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          sheetName: { type: 'string' },
          startCell: { type: 'string', description: 'Top-left cell, e.g. "A1"' },
          rows: {
            type: 'array',
            items: { type: 'array', items: { type: ['string', 'number'] } },
            description: 'Array of rows, each an array of cell values, e.g. [["Name","Total"],["Alice",100]]'
          }
        },
        required: ['fileName', 'sheetName', 'startCell', 'rows']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'read_data',
      description: 'Read cell values from a range, e.g. "A1:C10".',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          sheetName: { type: 'string' },
          range: { type: 'string' }
        },
        required: ['fileName', 'sheetName', 'range']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'apply_formula',
      description: 'Set a formula in a single cell, e.g. "=SUM(B2:B10)".',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          sheetName: { type: 'string' },
          cell: { type: 'string' },
          formula: { type: 'string' }
        },
        required: ['fileName', 'sheetName', 'cell', 'formula']
      }
    }
  }
];

export const EXCEL_TOOL_IMPL = {
  create_workbook: createWorkbook,
  list_sheets: listSheets,
  add_sheet: addSheet,
  write_data: writeData,
  read_data: readData,
  apply_formula: applyFormula
};
