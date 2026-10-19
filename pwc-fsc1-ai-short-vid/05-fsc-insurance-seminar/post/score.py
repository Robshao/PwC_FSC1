# 研討會短片暫用配樂（程式合成，無版權問題）：緊張 → 疑問 → 轉折 → 有信心 → 收尾
# 用法：python3 score.py 輸出.wav 總秒數
import numpy as np, wave, sys
SR = 48000
T = float(sys.argv[2])
out = np.zeros(int(SR * (T + 4)))
hz = lambda n: 440 * 2 ** ((n - 69) / 12)

def add(t0, w, vel):
    s = int(t0 * SR); out[s:s + len(w)] += vel * w

def piano(t0, note, vel=0.2, dur=2.5):
    t = np.arange(int(dur * SR)) / SR; f = hz(note)
    w = sum((1 / k ** 1.6) * np.sin(2 * np.pi * f * k * t * (1 + 0.0004 * k)) for k in range(1, 7))
    add(t0, w * np.minimum(t / 0.006, 1) * np.exp(-t * (1.6 + note / 60)), vel)

def pad(t0, t1, notes, vel=0.05, att=1.0, rel=1.0):
    t = np.arange(int((t1 - t0) * SR)) / SR; w = np.zeros(len(t))
    for m in notes:
        for d in (-0.003, 0, 0.003):
            f = hz(m) * (1 + d); w += np.sin(2 * np.pi * f * t) + 0.3 * np.sin(2 * np.pi * 2 * f * t)
    add(t0, w * np.minimum(t / att, 1) * np.minimum((t[-1] - t) / rel, 1) / len(notes), vel)

def pulse(t0, note, vel=0.3, dur=0.35):  # 低音脈動（像心跳／時鐘）
    t = np.arange(int(dur * SR)) / SR
    add(t0, np.sin(2 * np.pi * hz(note) * t) * np.exp(-t * 9) * np.minimum(t / 0.004, 1), vel)

def hit(t0, vel=0.6):  # 低頻衝擊（警示）
    t = np.arange(int(2.5 * SR)) / SR
    f = 55 * np.exp(-t * 1.5) + 30
    w = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t * 1.8)
    w += 0.15 * np.random.default_rng(1).standard_normal(len(t)) * np.exp(-t * 12)
    add(t0, w, vel)

def riser(t0, t1, vel=0.05):  # 上升音（帶入轉折）
    t = np.arange(int((t1 - t0) * SR)) / SR; p = t / t[-1]
    f = 200 + 1400 * p ** 2
    w = np.sin(2 * np.pi * np.cumsum(f) / SR) + 0.4 * np.random.default_rng(2).standard_normal(len(t)) * p
    add(t0, w * p ** 2, vel)

# 段落時間點（與 build.py 的剪接表一致）
CH, A4, B4, C4, D4, A5, B5, A6, B6, C6, SV, END = [float(x) for x in sys.argv[3].split(',')] if len(sys.argv) > 3 else \
    (0, 2.5, 5.5, 10.5, 13.5, 16.0, 24.0, 28.5, 31.8, 39.8, 44.3, 53.3)

# 一、問題浮現（D 小調）：低音持續音＋規律脈動
pad(0.2, C4 + 0.2, [38, 45, 50], 0.05, att=1.5, rel=0.4)
for tb in np.arange(A4, C4, 0.6):
    pulse(tb, 38, 0.22 + 0.06 * (tb - A4) / (C4 - A4))
for k, tb in enumerate(np.arange(A4 + 0.3, C4, 2.4)):
    piano(tb, [62, 65, 64, 62][k % 4], 0.07, 2.2)
# 二、警示：螢幕出現「觸發理賠」
hit(C4, 0.55)
# 三、疑問與緊張（Bb 增加不安感），心跳在問題卡時加快
n = int((A6 - C4) * SR); tt = np.arange(n) / SR
drone = np.sin(2 * np.pi * hz(33) * tt) + 0.6 * np.sin(2 * np.pi * hz(40) * tt) + 0.3 * np.sin(2 * np.pi * hz(46) * tt)
drone *= (0.6 + 0.4 * np.sin(2 * np.pi * 0.4 * tt)) * np.minimum(tt / 1.5, 1) * (0.5 + 0.5 * tt / tt[-1]) * np.minimum((tt[-1] - tt) / 0.3, 1)
add(C4, drone, 0.06)
for k, tb in enumerate(np.arange(C4 + 0.8, B5, 1.6)):
    piano(tb, [74, 75, 74, 77][k % 4], 0.06 + 0.004 * k, 2.0)
for tb in np.arange(B5, A6 - 0.4, 0.5):
    pulse(tb, 33, 0.35); pulse(tb + 0.16, 33, 0.18)
riser(A6 - 2.0, A6 - 0.05, 0.05)
# 四、轉折與解方（C 大調）：明亮的鋼琴琶音，逐漸增強
prog = [[48, 52, 55, 60, 64, 67], [43, 50, 55, 59, 62, 67], [45, 52, 57, 60, 64, 69], [41, 48, 53, 57, 60, 65]]
t = A6; i = 0
while t < END - 0.2:
    ch = prog[i % 4]; pad(t, min(t + 3.3, END + 0.5), ch[:3], 0.045 + 0.003 * i)
    for j, nt in enumerate([ch[3], ch[4], ch[5], ch[4], ch[3], ch[4]]):
        if t + j * 0.5 < END - 0.1:
            piano(t + j * 0.5, nt + (12 if t >= SV and j % 2 == 0 else 0), 0.1 + 0.008 * i, 2.5)
    if t >= SV:
        for b in range(6): pulse(t + b * 0.5, 36, 0.12)
    t += 3.0; i += 1
# 五、結尾：解決和弦慢慢消失
pad(END, T, [48, 52, 55, 60, 64], 0.09, att=0.4, rel=2.5)
for k, nt in enumerate([60, 72, 76, 79, 84]):
    piano(END + 0.1 + 0.1 * k, nt, 0.2 - 0.02 * k, 3.0)

out = out[:int(SR * T)]
out = out / np.max(np.abs(out)) * 0.7
st = np.stack([out, np.roll(out, int(0.012 * SR)) * 0.92], 1)
with wave.open(sys.argv[1], 'wb') as w:
    w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes((st * 32767).astype(np.int16).tobytes())
