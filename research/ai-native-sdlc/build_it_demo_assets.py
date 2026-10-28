# IT 現場示範素材：從 IT 團隊交付的任務包與錄影，產生提案簡報第「現場示範」頁與附錄用的圖片與影片
# 文件截圖是原文節錄（不改寫內容），用 pandoc 轉成網頁後以 Playwright 截圖；系統畫面取自錄影的單一影格
# 用法：python3 build_it_demo_assets.py <TASK-000013 解壓資料夾> <錄影資料夾> <輸出資料夾>
#   錄影資料夾內需有檔名包含 countries、regions、currencies 的 mp4
# 圖片與影片不進版本控制（見 .gitignore），建置簡報時以 --demo <輸出資料夾> 帶入
import sys, os, re, glob, shutil, subprocess, json, html

TASK, VID, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(OUT, exist_ok=True)
SPEC = os.path.join(TASK, 'specs')

def read(path):
    with open(path, encoding='utf-8') as f:
        return f.read()

def one(pattern):
    return glob.glob(os.path.join(SPEC, pattern))[0]

def rows(md, keys, cols=None):
    # 取出表格標頭與指定代碼開頭的列；cols 指定保留的欄位
    lines = [l for l in md.splitlines() if l.startswith('|')]
    head, sep = lines[0], lines[1]
    pick = [l for l in lines[2:] if any(f'**{k}' in l for k in keys)]
    def cut(l):
        cells = l.strip().strip('|').split('|')
        return '|' + '|'.join(cells[i] for i in cols) + '|' if cols else l
    return '\n'.join(cut(l) for l in [head, sep] + pick)

def section(md, start, end=None):
    a = md.index(start)
    b = md.index(end, a + len(start)) if end else len(md)
    return md[a:b].strip()

sa, qa = read(os.path.join(SPEC, 'SA_SPEC.md')), read(one('QA_SPEC/*.md'))
sd = read(one('SD_SPEC/*.md'))
t1 = read(glob.glob(os.path.join(TASK, 'TASK-*-001_*.md'))[0])
meta = re.search(r'governance_metadata:.*?(?=-->)', sa, re.S).group(0).strip()
tree = sorted(os.path.relpath(p, TASK) for p in glob.glob(os.path.join(SPEC, '**', '*.md'), recursive=True))
audit = '\n'.join(l.strip() for l in sd.splitlines() if 'Audit' in l and ('->>' in l or 'participant' in l))

# 五個步驟各一張：(檔名, 顯示的檔案路徑, 節錄的 Markdown)；框很小，只留重點、去掉路徑
base = lambda t: re.sub(r'`[^`]*/([^/`]+)`', r'`\1`', t)
trace = '\n'.join(re.sub(r'：`[^`]*`\s*', '：', l) for l in section(t1, '- **規格追溯依據**', '- **業務背景**').splitlines()[1:] if '架構標準' not in l)
STEPS = [
    ('step1.png', 'SA_SPEC.md',
     '## 業務規則（BR-0101-03）\n\n' + rows(sa, ['BR-FI-0'], [0, 1])),
    ('step2.png', 'specs/',
     '```\n' + '\n'.join(os.path.basename(p) for p in tree) + '\n```\n\n```\n' + '\n'.join(meta.splitlines()[1:4]) + '\n```'),
    ('step3.png', os.path.basename(glob.glob(os.path.join(TASK, 'TASK-*-001_*.md'))[0]),
     section(t1, '- **場景 3', '- **場景 4') + '\n\n**可修改檔案（最多 5 個）**\n\n' + base(section(t1, '1. `src/'))),
    ('step4.png', 'QA_SPEC.md',
     rows(qa, ['TC-FI-03', 'TC-FI-04', 'TC-FI-05', 'TC-FI-06', 'TC-FI-07'], [0, 1]) + '\n\n' + section(qa, '1. 後端')),
    ('step5.png', 'TASK-000013-001 · SD_SPEC',
     '**規格追溯依據**\n\n' + trace + '\n\n```\n' + audit.splitlines()[-1] + '\n```'),
]

CSS = '''
body{margin:0;font-family:"WenQuanYi Zen Hei","Microsoft JhengHei",sans-serif;font-size:17px;color:#222;background:#fff;overflow:hidden}
.bar{background:#2d2d2d;color:#eee;font:14px/1 monospace;padding:9px 12px}
.doc{padding:4px 14px}
h1,h2{font-size:18px;margin:10px 0 6px}
p{margin:6px 0}
table{border-collapse:collapse;width:100%;font-size:15px}
th,td{border:1px solid #ccc;padding:3px 6px;text-align:left;vertical-align:top}
th{background:#f2f2f2}
code,pre{font:13.5px/1.4 monospace;background:#f6f6f6}
pre{padding:8px 10px;white-space:pre-wrap}
ul{padding-left:20px;margin:6px 0}
li{margin:2px 0}
'''

pages = []
for name, title, md in STEPS:
    body = subprocess.run(['pandoc', '-f', 'gfm', '-t', 'html'], input=md, capture_output=True, text=True, check=True).stdout
    page = os.path.join(OUT, name.replace('.png', '.html'))
    with open(page, 'w', encoding='utf-8') as f:
        f.write(f'<!doctype html><meta charset="utf-8"><style>{CSS}</style><div class="bar">{html.escape(title)}</div><div class="doc">{body}</div>')
    pages.append((page, os.path.join(OUT, name)))

JS = '''
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 540, height: 395 }, deviceScaleFactor: 2 });
  for (const [src, dst] of JSON.parse(process.argv[2])) {
    await p.goto('file://' + src);
    await p.screenshot({ path: dst });
  }
  await b.close();
})();
'''
js = os.path.join(OUT, 'shot.js')
with open(js, 'w') as f: f.write(JS)
subprocess.run(['node', js, json.dumps([[os.path.abspath(a), os.path.abspath(b)] for a, b in pages])], check=True)
for a, _ in pages: os.remove(a)
os.remove(js)

# 系統畫面：各取一個資料乾淨的影格；影片複製一份供內部版嵌入
FRAMES = {'currencies': 3.0, 'countries': 1.5, 'regions': 4.5}
for key, t in FRAMES.items():
    src = glob.glob(os.path.join(VID, f'*{key}*.mp4'))[0]
    subprocess.run(['ffmpeg', '-v', 'error', '-y', '-ss', str(t), '-i', src, '-frames:v', '1', os.path.join(OUT, f'screen_{key}.png')], check=True)
    shutil.copy(src, os.path.join(OUT, f'video_{key}.mp4'))
print('assets in', OUT)
