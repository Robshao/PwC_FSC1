# 把手寫內容「寫」到白板上：用磁鐵當定位點逐格追蹤白板，再依台詞時間逐行顯示，人和手會自然擋在字前面
# 用法：python3 board_composite.py <鏡頭 6a|6b|6c> <版本 A|B> <輸入.mp4> <輸出.mp4>
import cv2, numpy as np, subprocess, sys, os
from magnets import detect

shot, V, src, dst = sys.argv[1:5]
INK = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'tools', 'remotion-cards', 'out', 'board')
W0, H0, W, H, FPS = 854, 480, 1920, 1080, 24
K = W / W0

# 基準座標（鏡頭 6b 的白板）上 8 顆磁鐵：名稱、x、y、顏色
CANON = {'tl': (434, 99, 'g'), 'c2': (519, 118, 'g'), 'b1': (588, 116, 'b'), 't4': (601, 87, 'g'),
         't5': (620, 92, 'g'), 'r6': (626, 157, 'g'), 'bl': (434, 224, 'g'), 'b2': (589, 222, 'b')}
# 各鏡頭第一格畫面中的對應位置（初始追蹤用）
START = {
    '6a': {'tl': (298, 80), 'c2': (436, 142), 'b1': (523, 144), 't4': (537, 85), 't5': (557, 99), 'r6': (561, 227), 'bl': (300, 418)},
    '6b': {k: v[:2] for k, v in CANON.items()},
    '6c': {'tl': (313, 105), 'c2': (374, 120), 'b1': (423, 118), 't4': (432, 96), 't5': (445, 100), 'r6': (450, 150), 'bl': (314, 201), 'b2': (424, 199)},
}[shot]
# 逐行出現的時間（這個鏡頭內的秒數）：(圖層, 開始, 結束)
REVEAL = {
    '6a': [('base', -1, -1)],
    '6b': {'B': [('base', -1, -1), ('i1', 2.55, 3.15), ('i2', 3.65, 4.25), ('i3', 4.85, 5.45)],
           'A': [('base', -1, -1), ('i1', 0.5, 1.4), ('i2', 2.9, 3.8), ('i3', 5.5, 6.4)]}[V],
    '6c': [('base', -1, -1), ('i1', -1, -1), ('i2', -1, -1), ('i3', -1, -1)],
}[shot]
RADIUS = 30 if shot == '6a' else 16

layers = {}
for name, _, _ in REVEAL:
    im = cv2.imread(f'{INK}/{V}-{name}.png', cv2.IMREAD_UNCHANGED).astype(np.float32)
    ys, xs = np.where(im[..., 3] > 0)
    layers[name] = (im, xs.min(), xs.max())

names = list(START)
cp = np.float32([CANON[n][:2] for n in names])
Hm, _ = cv2.findHomography(cp, np.float32([START[n] for n in names]))

# 1. 逐格追蹤
frames = []
p = subprocess.Popen(['ffmpeg', '-v', 'error', '-i', src, '-vf', f'fps={FPS}', '-f', 'rawvideo', '-pix_fmt', 'bgr24', '-'], stdout=subprocess.PIPE)
while True:
    b = p.stdout.read(W0 * H0 * 3)
    if len(b) < W0 * H0 * 3: break
    frames.append(np.frombuffer(b, np.uint8).reshape(H0, W0, 3).copy())
Hs = []
allc = np.float32([[c[0], c[1]] for c in CANON.values()])
for f in frames:
    det = detect(f)
    pred = cv2.perspectiveTransform(allc.reshape(-1, 1, 2), Hm).reshape(-1, 2)
    src_pts, dst_pts = [], []
    for (cx, cy, col), (px, py) in zip(CANON.values(), pred):
        cands = [(np.hypot(x - px, y - py), x, y) for x, y, t, a in det if t == col]
        if cands:
            d, x, y = min(cands)
            if d < RADIUS: src_pts.append((cx, cy)); dst_pts.append((x, y))
    if len(src_pts) >= 5:
        Hn, inl = cv2.findHomography(np.float32(src_pts), np.float32(dst_pts), cv2.RANSAC, 3.0)
        if Hn is not None and inl.sum() >= 4: Hm = Hn
    Hs.append(Hm.copy())
# 平滑：用白板四角的軌跡做移動平均
box = np.float32([[300, 40], [785, 40], [785, 330], [300, 330]]).reshape(-1, 1, 2)
corners = np.array([cv2.perspectiveTransform(box, h).reshape(-1, 2) for h in Hs])
sm = np.copy(corners)
for i in range(len(corners)):
    sm[i] = corners[max(0, i - 3):i + 4].mean(0)
S = np.diag([K, K, 1.0])

# 2. 合成
out = subprocess.Popen(['ffmpeg', '-v', 'error', '-y', '-f', 'rawvideo', '-pix_fmt', 'bgr24', '-s', f'{W}x{H}', '-r', str(FPS), '-i', '-',
                        '-i', src, '-map', '0:v', '-map', '1:a', '-c:v', 'libx264', '-crf', '15', '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '192k', '-shortest', dst],
                       stdin=subprocess.PIPE)
for i, f in enumerate(frames):
    t = i / FPS
    h = cv2.getPerspectiveTransform(box.reshape(-1, 2), sm[i].astype(np.float32))
    h1080 = S @ h @ np.linalg.inv(S)
    big = cv2.resize(f, (W, H), interpolation=cv2.INTER_LANCZOS4).astype(np.float32)
    ink = np.zeros((H, W, 4), np.float32)
    for name, t0, t1 in REVEAL:
        im, x0, x1 = layers[name]
        if t0 >= 0:
            if t < t0: continue
            prog = min(1.0, (t - t0) / (t1 - t0))
            cut = int(x0 + (x1 - x0) * prog)
            im = im.copy(); im[:, cut + 1:, 3] = 0
        a = im[..., 3:4] / 255
        ink[..., :3] = ink[..., :3] * (1 - a) + im[..., :3] * a
        ink[..., 3:4] = np.maximum(ink[..., 3:4], a)
    wk = cv2.warpPerspective(ink, h1080, (W, H), flags=cv2.INTER_LINEAR)
    wk = cv2.GaussianBlur(wk, (0, 0), 0.9)
    # 只畫在白板上：低彩度、不太暗（陰影裡的白板也算）；皮膚、西裝會被排除
    hsv = cv2.cvtColor(big.astype(np.uint8), cv2.COLOR_BGR2HSV)
    board = ((hsv[..., 1] < 50) & (hsv[..., 2] > 95)).astype(np.uint8)
    board = cv2.erode(board, np.ones((5, 5), np.uint8))
    board = cv2.GaussianBlur(board.astype(np.float32), (0, 0), 1.5)[..., None]
    a = wk[..., 3:4] * 0.92 * board
    # 正片疊底：墨水會跟著白板的明暗（陰影處的字也會變暗）
    res = big * (1 - a + a * wk[..., :3] / 255)
    out.stdin.write(np.clip(res, 0, 255).astype(np.uint8).tobytes())
out.stdin.close(); out.wait()
print(dst, len(frames), 'frames')
