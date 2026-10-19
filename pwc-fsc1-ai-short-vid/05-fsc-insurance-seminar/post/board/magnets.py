# 偵測白板上的綠色、藍色圓形磁鐵（白板追蹤用的定位點）
import cv2, numpy as np

def detect(img):
    """img: BGR 854x480。回傳 [(x, y, 'g'|'b', 面積)]"""
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    out = []
    for tag, lo, hi in (('g', (55, 90, 50), (90, 255, 230)), ('b', (95, 120, 60), (125, 255, 255))):
        m = cv2.inRange(hsv, np.array(lo), np.array(hi))
        m = cv2.morphologyEx(m, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))
        n, lab, st, cen = cv2.connectedComponentsWithStats(m)
        for i in range(1, n):
            a = st[i, cv2.CC_STAT_AREA]; w = st[i, cv2.CC_STAT_WIDTH]; h = st[i, cv2.CC_STAT_HEIGHT]
            if 12 <= a <= 3000 and 0.5 < w / max(h, 1) < 2.0:
                out.append((float(cen[i][0]), float(cen[i][1]), tag, int(a)))
    return out
