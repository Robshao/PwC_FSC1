import cv2, numpy as np, glob, json, sys
S='/tmp/claude-0/-home-user-PwC-FSC1/bfb60935-eecd-590b-8578-281e4d2840a2/scratchpad/v5/tab'
fs=sorted(glob.glob(S+'/fr/*.png')); Q=[np.float32(q) for q in json.load(open(S+'/quads_raw.json'))]
# temporal smoothing of corners (5-frame moving average) to remove jitter
A=np.stack(Q); K=5; P=np.pad(A,((K//2,K//2),(0,0),(0,0)),mode='edge')
A=np.stack([P[i:i+K].mean(0) for i in range(len(Q))])
dash=cv2.imread('/tmp/claude-0/-home-user-PwC-FSC1/bfb60935-eecd-590b-8578-281e4d2840a2/images/26.png')[46:722,46:980]
dash=cv2.resize(dash,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC).astype(np.float32)
dash*=np.float32([0.88,0.92,0.97])  # warm afternoon tint (BGR)
h,w=dash.shape[:2]; src=np.float32([[0,0],[w,0],[w,h],[0,h]])
import os; os.makedirs(S+'/out',exist_ok=True)
only=[int(x) for x in sys.argv[1:]] if len(sys.argv)>1 else range(len(fs))
for i in only:
    im=cv2.imread(fs[i]).astype(np.float32); q=A[i]
    H=cv2.getPerspectiveTransform(src,q)
    wd=cv2.warpPerspective(dash,H,(1280,720),flags=cv2.INTER_LINEAR)
    wd=cv2.GaussianBlur(wd,(0,0),0.7)
    qm=np.zeros((720,1280),np.uint8); cv2.fillConvexPoly(qm,np.round(q).astype(np.int32),255)
    qm=cv2.erode(qm,np.ones((3,3),np.uint8))
    hsv=cv2.cvtColor(im.astype(np.uint8),cv2.COLOR_BGR2HSV)
    Hh,Ss,Vv=hsv[...,0],hsv[...,1],hsv[...,2]
    skin=((Hh>=6)&(Hh<=24)&(Ss>=35)&(Ss<=175)&(Vv>=70)).astype(np.uint8)*255
    skin=cv2.morphologyEx(skin,cv2.MORPH_OPEN,np.ones((5,5),np.uint8))
    skin=cv2.morphologyEx(skin,cv2.MORPH_CLOSE,np.ones((9,9),np.uint8))
    # keep only large skin blobs (hand / sleeve), drop specks from orange dashes
    n,lab,st,_=cv2.connectedComponentsWithStats(skin)
    keep=np.zeros_like(skin)
    for k in range(1,n):
        if st[k,4]>400: keep[lab==k]=255
    keep=cv2.dilate(keep,np.ones((3,3),np.uint8))
    alpha=(qm>0).astype(np.float32)*(1-(keep>0).astype(np.float32))
    alpha=cv2.GaussianBlur(alpha,(0,0),1.2)[...,None]
    # subtle screen glare gradient to blend
    out=im*(1-alpha)+wd*alpha
    cv2.imwrite(S+f'/out/{i+1:03d}.png',np.clip(out,0,255).astype(np.uint8))
