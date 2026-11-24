// JSON Lines client for the Lean Test Oracle
// (lean_spec/.lake/build/bin/bloggas_oracle). One request line in,
// one response line out, strictly in order.
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

export const ORACLE_BIN = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '..', 'lean_spec', '.lake', 'build', 'bin', 'bloggas_oracle',
);

export const startOracle = () => {
  const proc = spawn(ORACLE_BIN, [], { stdio: ['pipe', 'pipe', 'inherit'] });
  const rl = createInterface({ input: proc.stdout });
  const waiters = [];
  rl.on('line', (line) => {
    const resolve = waiters.shift();
    if (resolve) resolve(JSON.parse(line));
  });
  const ask = (request) => new Promise((resolve, reject) => {
    waiters.push(resolve);
    proc.stdin.write(`${JSON.stringify(request)}\n`, (err) => {
      if (err) reject(err);
    });
  });
  const stop = () => {
    rl.close();
    proc.stdin.end();
    proc.kill();
  };
  return { ask, stop };
};
