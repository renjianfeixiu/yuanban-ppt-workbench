import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';
import { homedir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { setTimeout as delay } from 'node:timers/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = path.join(root, '启动产品Demo.command');
const modes = [
  ['直接运行启动文件', [entry]],
  ['按文本执行启动命令', ['-fc', 'eval "$(<"$1")"', 'zsh', entry]],
  ['source 启动文件且保留调用方的零号参数', ['-fc', 'unsetopt FUNCTION_ARGZERO; source "$1"', 'zsh', entry]],
];

async function unusedPort() {
  const socket = createServer();
  await new Promise((resolve, reject) => { socket.once('error', reject); socket.listen(0, '127.0.0.1', resolve); });
  const port = socket.address().port;
  await new Promise(resolve => socket.close(resolve));
  return port;
}

for (const [name, args] of modes) {
  test(`${name}：从用户目录启动并实际提供工作台`, { timeout: 50000 }, async () => {
    const port = await unusedPort();
    const processUnderTest = spawn('/bin/zsh', args, {
      cwd: homedir(), detached: true,
      env: { ...process.env, PPT_DEMO_PORT: String(port), PPT_DEMO_NO_BROWSER: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let output = '', exit;
    processUnderTest.stdout.on('data', x => { output = (output + x).slice(-12000); });
    processUnderTest.stderr.on('data', x => { output = (output + x).slice(-12000); });
    const closed = new Promise(resolve => processUnderTest.once('exit', (code, signal) => { exit = { code, signal }; resolve(); }));
    try {
      const base = `http://127.0.0.1:${port}`;
      let ready = false;
      const deadline = Date.now() + 40000;
      while (Date.now() < deadline) {
        if (exit) throw new Error(`启动入口提前退出 ${JSON.stringify(exit)}\n${output}`);
        try {
          const response = await fetch(base + '/api/bootstrap', { signal: AbortSignal.timeout(800) });
          const bootstrap = await response.json();
          ready = response.ok && bootstrap.template?.slideCount === 22 && bootstrap.pages?.length === 22;
        } catch { /* The template is still loading. */ }
        if (ready) break;
        await delay(200);
      }
      assert.ok(ready, `启动后没有可用服务\n${output}`);
      const page = await fetch(base + '/prototype/ppt-template');
      assert.equal(page.status, 200);
      assert.match(await page.text(), /id="app"/);
    } finally {
      if (!exit) process.kill(-processUnderTest.pid, 'SIGTERM');
      await Promise.race([closed, delay(3000)]);
      if (!exit) { process.kill(-processUnderTest.pid, 'SIGKILL'); await closed; }
    }
  });
}
