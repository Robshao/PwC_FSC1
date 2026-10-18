// 用法：node render.mjs <字卡名稱> <props.json> [輸出檔]  [--overlay]
// --overlay：輸出透明背景的 ProRes 4444 .mov，用來疊在影片上（問題卡用）
import {spawnSync} from 'node:child_process';
import {existsSync, mkdirSync} from 'node:fs';

const args = process.argv.slice(2);
const flags = args.filter((a) => a.startsWith('--'));
const [id, props, outArg] = args.filter((a) => !a.startsWith('--'));
if (!id || !props) {
  console.error('用法：node render.mjs <ChapterCard|QuestionCards|ServiceCard|EndCard> <props.json> [輸出檔]  [--overlay]');
  process.exit(1);
}
const overlay = flags.includes('--overlay');
mkdirSync('out', {recursive: true});
const out = outArg ?? `out/${id}${overlay ? '.mov' : '.mp4'}`;

const browser = [process.env.REMOTION_BROWSER, '/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell'].find((p) => p && existsSync(p));
const cmd = ['remotion', 'render', 'src/index.ts', id, out, `--props=${props}`];
if (browser) cmd.push(`--browser-executable=${browser}`);
if (overlay) cmd.push('--codec=prores', '--prores-profile=4444', '--image-format=png', '--pixel-format=yuva444p10le');

const r = spawnSync('npx', cmd, {stdio: 'inherit'});
process.exit(r.status ?? 1);
