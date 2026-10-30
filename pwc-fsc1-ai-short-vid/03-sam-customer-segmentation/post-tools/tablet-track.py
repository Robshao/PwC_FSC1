import cv2, numpy as np, glob, json
S='/tmp/claude-0/-home-user-PwC-FSC1/bfb60935-eecd-590b-8578-281e4d2840a2/scratchpad/v5/tab'
fs=sorted(glob.glob(S+'/fr/*.png'))
Q0=np.float32([[658,420],[850,404],[905,555],[710,582]])
sift=cv2.SIFT_create(4000)
def feat(im,quad):
    m=np.zeros(im.shape[:2],np.uint8); cv2.fillConvexPoly(m,quad.astype(np.int32),255)
    m=cv2.dilate(m,np.ones((25,25),np.uint8))
    g=cv2.cvtColor(im,cv2.COLOR_BGR2GRAY); g=cv2.createCLAHE(2.0,(8,8)).apply(g)
    return sift.detectAndCompute(g,m)
bf=cv2.BFMatcher()
quads=[Q0]; prev=cv2.imread(fs[0]); kp0,d0=feat(prev,Q0); prevq=Q0
for i,f in enumerate(fs[1:],1):
    im=cv2.imread(f); kp,d=feat(im,prevq)
    # match to previous frame (chain) for robustness
    kpp,dp=feat(prev,prevq)
    ms=bf.knnMatch(dp,d,k=2); good=[a for a,b in ms if a.distance<0.75*b.distance]
    src=np.float32([kpp[m.queryIdx].pt for m in good]); dst=np.float32([kp[m.trainIdx].pt for m in good])
    H,inl=cv2.findHomography(src,dst,cv2.RANSAC,2.0)
    q=cv2.perspectiveTransform(prevq.reshape(-1,1,2),H).reshape(-1,2)
    quads.append(q); prev=im; prevq=q
    if i%15==0: print(i,len(good),int(inl.sum()),np.round(q).astype(int).tolist())
json.dump([q.tolist() for q in quads],open(S+'/quads_raw.json','w'))
for n in (0,45,89):
    im=cv2.imread(fs[n]); cv2.polylines(im,[quads[n].astype(np.int32)],True,(0,255,0),2)
    cv2.imwrite(S+f'/q{n}.png',im[330:720,560:1000])
