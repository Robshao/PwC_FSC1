# 研討會短片：片名卡＋鏡頭 4–6（白板手寫內容）＋服務卡＋結尾卡
# 用法：python3 build.py A|B   （A＝原版台詞、B＝新版台詞）
import subprocess, sys, os, json
V = sys.argv[1] if len(sys.argv) > 1 else 'B'
HERE = os.path.dirname(os.path.abspath(__file__))
CLIPS = os.path.join(HERE, '..', 'clips')
CARDS = os.path.join(HERE, '..', '..', 'tools', 'remotion-cards', 'out')
OUT = os.path.join(HERE, '..', 'output'); os.makedirs(OUT, exist_ok=True)
TMP = os.path.join(OUT, f'tmp-{V}'); os.makedirs(TMP, exist_ok=True)
FONT = 'WenQuanYi Zen Hei'

# (名稱, 來源, 入點, 出點, 聲音類型)  聲音類型：dlg＝台詞、amb＝環境音、none＝無聲
EDIT = [
    ('chapter', f'{CARDS}/TitleCard.mp4', 0, 3.2, 'none'),
    ('4a', f'{CLIPS}/s4a.mp4', 0, 3.0, 'amb'),
    ('4b', f'{CLIPS}/s4b.mp4', 0, 5.0, 'dlg'),
    ('4c', f'{CLIPS}/s4c.mp4', 0.3, 3.3, 'amb'),
    ('4d', f'{CLIPS}/s4d.mp4', 0, 2.5, 'amb'),
    # 5a＋5b 合成一段：怡君說到一半切到黃副總（J-cut），問題卡跟著她列舉的項目浮現
    ('5ab', None, 0, 9.2, 'special'),
    ('6a', f'{CLIPS}/s6a-ink-{V}.mp4', 0.3, 3.6, 'amb'),
    ('6b', f'{CLIPS}/s6b-{V}-ink-{V}.mp4', 0, 8.0, 'dlg'),
    ('6c', f'{CLIPS}/s6c-ink-{V}.mp4', 0.3, 4.8, 'amb'),
    ('service', f'{CARDS}/ServiceCard.mp4', 0, 9.0, 'none'),
    ('end', f'{CARDS}/EndCard.mp4', 0, 6.0, 'none'),
]
J5A = 4.5          # 5a 畫面用到第幾秒就切到 5b
FADE_OUT = {'6c': 0.5}   # 這段結尾淡出到黑
FADE_IN = {'service': 0.4}
start = {}; t = 0.0
for name, _, a, b, _ in EDIT:
    start[name] = t; t += b - a
TOTAL = t

def run(cmd):
    subprocess.run(cmd, check=True)

# 1. 每段統一成 1920x1080、30fps；聲音依類型調整音量
parts = []
for name, src, a, b, kind in EDIT:
    d = b - a; o = f'{TMP}/{name}.mp4'
    vf = 'scale=1920:1080:flags=lanczos,unsharp=5:5:0.6,fps=30,format=yuv420p'
    if name in FADE_OUT: vf += f',fade=t=out:st={d-FADE_OUT[name]:.2f}:d={FADE_OUT[name]}'
    if name in FADE_IN: vf += f',fade=t=in:st=0:d={FADE_IN[name]}'
    if kind == 'special':
        a5, b5 = f'{CLIPS}/s5a.mp4', f'{CLIPS}/s5b.mp4'
        tail = d - 8.0; v5b = d - J5A
        fc = (f"[0:v]trim=0:{J5A},setpts=PTS-STARTPTS,{vf}[v1];[1:v]trim=0.2:{0.2+v5b},setpts=PTS-STARTPTS,{vf}[v2];[v1][v2]concat=n=2:v=1:a=0[v];"
              f"[0:a]atrim=0:8.0,asetpts=PTS-STARTPTS,loudnorm=I=-16:TP=-2:LRA=11,aresample=48000,afade=t=out:st=7.9:d=0.1[x1];"
              f"[1:a]atrim={0.2+v5b-tail}:{0.2+v5b},asetpts=PTS-STARTPTS,loudnorm=I=-30:TP=-2:LRA=11,aresample=48000,afade=t=in:d=0.1[x2];"
              f"[x1][x2]concat=n=2:v=0:a=1[a]")
        run(['ffmpeg', '-v', 'error', '-y', '-i', a5, '-i', b5, '-filter_complex', fc, '-map', '[v]', '-map', '[a]', '-ac', '2',
             '-c:v', 'libx264', '-crf', '16', '-preset', 'medium', '-c:a', 'pcm_s16le', '-ar', '48000', o.replace('.mp4', '.mov')])
        parts.append(o.replace('.mp4', '.mov')); continue
    if kind == 'none':
        cmd = ['ffmpeg', '-v', 'error', '-y', '-ss', str(a), '-t', str(d), '-i', src,
               '-f', 'lavfi', '-t', str(d), '-i', 'anullsrc=r=48000:cl=stereo',
               '-map', '0:v', '-map', '1:a', '-vf', vf]
    else:
        lvl = '-16' if kind == 'dlg' else '-30'
        af = f'loudnorm=I={lvl}:TP=-2:LRA=11,aresample=48000,afade=t=in:d=0.03,afade=t=out:st={d-0.04:.2f}:d=0.04'
        cmd = ['ffmpeg', '-v', 'error', '-y', '-ss', str(a), '-t', str(d), '-i', src, '-vf', vf, '-af', af, '-ac', '2']
    run(cmd + ['-c:v', 'libx264', '-crf', '16', '-preset', 'medium', '-c:a', 'pcm_s16le', '-ar', '48000', o.replace('.mp4', '.mov')])
    parts.append(o.replace('.mp4', '.mov'))
with open(f'{TMP}/list.txt', 'w') as f:
    for p in parts: f.write(f"file '{p}'\n")
run(['ffmpeg', '-v', 'error', '-y', '-f', 'concat', '-safe', '0', '-i', f'{TMP}/list.txt', '-c', 'copy', f'{TMP}/cut.mov'])

# 2. 字幕與標籤（ASS）
S = start
def ts(x):
    h = int(x // 3600); m = int(x % 3600 // 60); s = x % 60
    return f'{h}:{m:02d}:{s:05.2f}'
subs = [
    (S['4a'] + 0.3, S['4a'] + 2.9, 'Caption', '市場上已出現 AI 專屬保險'),
    (S['4b'] + 0.2, S['4b'] + 3.0, 'Name', '黃副總｜產險公司商品部'),
    (S['5ab'] + 0.2, S['5ab'] + 3.0, 'Name', '怡君｜核保主管'),
    (S['6b'] + 0.1, S['6b'] + 3.2, 'Name', '欣怡｜FSC1 保險組顧問'),
    (S['4b'] + 0.3, S['4b'] + 2.6, 'Sub', 'AI 沒達到約定的表現，'),
    (S['4b'] + 2.7, S['4b'] + 5.0, 'Sub', '損失就由保險來承擔。'),
    (S['5ab'] + 0.2, S['5ab'] + 3.2, 'Sub', '再保公司給了商品，'),
    (S['5ab'] + 3.3, S['5ab'] + 8.0, 'Sub', '可是系統、數據、模型評估，我們接得住嗎？'),
]
if V == 'A':
    subs += [(S['6b'] + 0.5, S['6b'] + 2.7, 'Sub', '先評估 AI 模型的風險，'),
             (S['6b'] + 2.8, S['6b'] + 5.3, 'Sub', '再把企業的使用情境分層，'),
             (S['6b'] + 5.4, S['6b'] + 8.0, 'Sub', '商品跟核保才有依據。')]
else:
    subs += [(S['6b'] + 0.1, S['6b'] + 2.4, 'Sub', '商品交給再保公司，'),
             (S['6b'] + 2.5, S['6b'] + 5.7, 'Sub', '系統、數據、治理，'),
             (S['6b'] + 5.8, S['6b'] + 8.0, 'Sub', '我們陪你一起接起來。')]
ass = f"""[Script Info]
ScriptType: v4.00+
PlayResX: 1920
PlayResY: 1080
WrapStyle: 2

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Sub,{FONT},56,&H00FFFFFF,&H00FFFFFF,&H00101010,&H80000000,1,0,0,0,100,100,1,0,1,3,1,2,80,80,70,1
Style: Caption,{FONT},50,&H00FFFFFF,&H00FFFFFF,&H002E1A0B,&H00000000,1,0,0,0,100,100,2,0,3,22,0,1,110,80,110,1
Style: Name,{FONT},38,&H00FFFFFF,&H00FFFFFF,&H000F48D9,&H00000000,1,0,0,0,100,100,2,0,3,14,0,1,90,80,190,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""
for a, b, st, txt in subs:
    ass += f'Dialogue: 0,{ts(a)},{ts(b)},{st},,0,0,0,,{{\\fad(200,200)}}{txt}\n'
open(f'{TMP}/subs.ass', 'w', encoding='utf8').write(ass)

# 3. 配樂（依剪接點合成）
M = dict(S); M['5a'] = S['5ab']; M['5b'] = S['5ab'] + J5A
marks = ','.join(f"{M[k]:.2f}" for k in ['chapter', '4a', '4b', '4c', '4d', '5a', '5b', '6a', '6b', '6c', 'service', 'end'])
run(['python3', os.path.join(HERE, 'score.py'), f'{TMP}/music.wav', f'{TOTAL:.2f}', marks])

# 4. 合成：問題卡疊在 5b、燒字幕、配樂在台詞時自動降低音量、頭尾淡入淡出
q0 = S['5ab'] + J5A
fc = (f"[1:v]setpts=PTS-STARTPTS+{q0}/TB[q];[0:v][q]overlay=format=auto:eof_action=pass,"
      f"subtitles={TMP}/subs.ass,fade=t=in:st=0:d=0.4,fade=t=out:st={TOTAL-1.2:.2f}:d=1.2[v];"
      f"[0:a]asplit=2[dlg][sc];[2:a]volume=0.55[mus];"
      f"[mus][sc]sidechaincompress=threshold=0.03:ratio=8:attack=20:release=400[duck];"
      f"[dlg][duck]amix=inputs=2:duration=first:normalize=0,afade=t=out:st={TOTAL-1.5:.2f}:d=1.5,loudnorm=I=-16:TP=-1.5:LRA=11[a]")
name = '原版' if V == 'A' else '新版'
final = f'{OUT}/seminar-4to6-draft-{name}.mp4'
run(['ffmpeg', '-v', 'error', '-y', '-i', f'{TMP}/cut.mov', '-i', f'{CARDS}/QuestionCards.mov', '-i', f'{TMP}/music.wav',
     '-filter_complex', fc, '-map', '[v]', '-map', '[a]', '-c:v', 'libx264', '-crf', '18', '-preset', 'medium',
     '-c:a', 'aac', '-b:a', '192k', '-ar', '48000', '-movflags', '+faststart', final])
print(final, f'{TOTAL:.1f}s')
