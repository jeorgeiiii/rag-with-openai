import { spawn } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';

const TIMEOUT_MS = 30_000;

/**
 * Runs one COM action via a PowerShell helper script (excel-com.ps1,
 * word-com.ps1, powerpoint-com.ps1, ...). Windows + the relevant Office app
 * only - by design this never runs on the Vercel deployment.
 */
export async function runComAction(scriptPath, appLabel, action, payload) {
  if (process.platform !== 'win32') {
    throw new Error(`${appLabel} automation requires Windows with Microsoft ${appLabel} installed (run this app locally, not on the server deployment).`);
  }

  const tmpDir = os.tmpdir();
  const stamp = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const payloadPath = path.join(tmpDir, `office-payload-${stamp}.json`);
  const resultPath = path.join(tmpDir, `office-result-${stamp}.json`);

  fs.writeFileSync(payloadPath, JSON.stringify(payload), 'utf8');

  try {
    await new Promise((resolve, reject) => {
      const child = spawn('powershell.exe', [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', scriptPath,
        '-Action', action,
        '-PayloadPath', payloadPath,
        '-ResultPath', resultPath
      ]);

      const timer = setTimeout(() => {
        child.kill();
        // The script writes its result before the (sometimes slow-to-exit)
        // app-quit/cleanup step in its finally block, so a still-running
        // process at this point doesn't mean the actual work failed - check
        // for a completed result file before treating this as a real timeout.
        if (fs.existsSync(resultPath)) {
          resolve();
        } else {
          reject(new Error(`${appLabel} automation timed out after ${TIMEOUT_MS}ms`));
        }
      }, TIMEOUT_MS);

      let stderr = '';
      child.stderr.on('data', (chunk) => { stderr += chunk.toString(); });

      child.on('error', (err) => {
        clearTimeout(timer);
        reject(err);
      });

      child.on('close', (code) => {
        clearTimeout(timer);
        if (code !== 0 && !fs.existsSync(resultPath)) {
          reject(new Error(stderr || `powershell exited with code ${code}`));
        } else {
          resolve();
        }
      });
    });

    if (!fs.existsSync(resultPath)) {
      throw new Error(`${appLabel} automation produced no result`);
    }

    let raw = fs.readFileSync(resultPath, 'utf8');
    if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1); // strip PowerShell's UTF-8 BOM
    const result = JSON.parse(raw);
    if (!result.ok) {
      throw new Error(result.error || `${appLabel} automation failed`);
    }
    return result;
  } finally {
    for (const p of [payloadPath, resultPath]) {
      try { fs.unlinkSync(p); } catch {}
    }
  }
}
