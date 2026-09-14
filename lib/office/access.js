import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { runComAction } from './com-runner.js';
import { resolveSafePath } from './workdir.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCRIPT_PATH = path.resolve(__dirname, '../../scripts/office/access-com.ps1');

const IDENTIFIER_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

function assertIdentifier(name, label) {
  if (!IDENTIFIER_RE.test(name || '')) {
    throw new Error(`Invalid ${label}: "${name}" (letters, digits, underscore only, starting with a letter or underscore)`);
  }
}

function runAccessAction(action, payload) {
  return runComAction(SCRIPT_PATH, 'Access', action, payload);
}

export async function createDatabase({ fileName }) {
  const filePath = resolveSafePath(fileName, 'accdb');
  if (fs.existsSync(filePath)) {
    throw new Error(`"${path.basename(filePath)}" already exists. Choose a different name.`);
  }
  return runAccessAction('CreateDatabase', { filePath });
}

export async function createTable({ fileName, tableName, columns }) {
  assertIdentifier(tableName, 'tableName');
  for (const col of columns || []) {
    assertIdentifier(col.name, 'column name');
  }
  return runAccessAction('CreateTable', {
    filePath: resolveSafePath(fileName, 'accdb'),
    tableName,
    columns: columns || []
  });
}

export async function insertRows({ fileName, tableName, rows }) {
  assertIdentifier(tableName, 'tableName');
  return runAccessAction('InsertRows', {
    filePath: resolveSafePath(fileName, 'accdb'),
    tableName,
    rows: rows || []
  });
}

export async function queryTable({ fileName, sql }) {
  if (!/^\s*select\b/i.test(sql || '')) {
    throw new Error('Only SELECT queries are allowed. Use insert_rows to add data.');
  }
  return runAccessAction('QueryTable', { filePath: resolveSafePath(fileName, 'accdb'), sql });
}

export async function listTables({ fileName }) {
  return runAccessAction('ListTables', { filePath: resolveSafePath(fileName, 'accdb') });
}

// OpenAI/Groq-compatible tool (function calling) definitions.
export const ACCESS_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'create_database',
      description: 'Create a new Access database (.accdb) in the chatbot office folder. Fails if the file already exists.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string', description: 'File name, e.g. "Contacts". .accdb is added automatically.' }
        },
        required: ['fileName']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'create_table',
      description: 'Create a new table in an Access database. An auto-numbering ID primary key column is added automatically - do not define one yourself.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          tableName: { type: 'string', description: 'Letters, digits, underscore only.' },
          columns: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                name: { type: 'string', description: 'Letters, digits, underscore only.' },
                type: { type: 'string', description: 'Access SQL type, e.g. "TEXT(255)", "LONG", "DOUBLE", "DATETIME", "YESNO".' }
              },
              required: ['name', 'type']
            }
          }
        },
        required: ['fileName', 'tableName', 'columns']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'insert_rows',
      description: 'Insert rows into an existing table.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          tableName: { type: 'string' },
          rows: {
            type: 'array',
            items: { type: 'object' },
            description: 'Array of objects keyed by column name, e.g. [{"Name":"Alice","Age":30}]'
          }
        },
        required: ['fileName', 'tableName', 'rows']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'query_table',
      description: 'Run a read-only SELECT query against the database and return matching rows.',
      parameters: {
        type: 'object',
        properties: {
          fileName: { type: 'string' },
          sql: { type: 'string', description: 'A SELECT statement, e.g. "SELECT * FROM People WHERE Age > 25"' }
        },
        required: ['fileName', 'sql']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'list_tables',
      description: 'List the user tables in an Access database.',
      parameters: {
        type: 'object',
        properties: { fileName: { type: 'string' } },
        required: ['fileName']
      }
    }
  }
];

export const ACCESS_TOOL_IMPL = {
  create_database: createDatabase,
  create_table: createTable,
  insert_rows: insertRows,
  query_table: queryTable,
  list_tables: listTables
};
