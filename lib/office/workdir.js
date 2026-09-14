import fs from 'fs';
import os from 'os';
import path from 'path';

// All Office automation is confined to this single folder so the chatbot can
// never read or write arbitrary files on the machine via a crafted file name.
const WORKDIR = process.env.OFFICE_WORKDIR
  ? path.resolve(process.env.OFFICE_WORKDIR)
  : path.join(os.homedir(), 'Documents', 'ChatbotOffice');

export function getWorkdir() {
  fs.mkdirSync(WORKDIR, { recursive: true });
  return WORKDIR;
}

/**
 * Resolves a user-supplied file name to an absolute path inside the office
 * working directory, rejecting anything that would escape it (../, absolute
 * paths, alternate drive letters, etc). Appends `extension` (e.g. "xlsx",
 * "docx", "pptx") when the name doesn't already end with it.
 */
export function resolveSafePath(fileName, extension) {
  if (!fileName || typeof fileName !== 'string') {
    throw new Error('fileName is required');
  }
  if (!extension || typeof extension !== 'string') {
    throw new Error('extension is required');
  }

  const base = fileName.trim().replace(/\\/g, '/').split('/').pop();
  if (!base || base === '.' || base === '..') {
    throw new Error('Invalid fileName');
  }

  const ext = extension.replace(/^\./, '').toLowerCase();
  const withExt = new RegExp(`\\.${ext}$`, 'i').test(base) ? base : `${base}.${ext}`;
  const workdir = getWorkdir();
  const resolved = path.resolve(workdir, withExt);

  if (!resolved.startsWith(workdir + path.sep)) {
    throw new Error('fileName escapes the office working directory');
  }

  return resolved;
}
