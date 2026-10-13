import numpy as np, wave, sys
SR=48000; T=45.0
out=np.zeros(int(SR*(T+4)))
def hz(n): return 440*2**((n-69)/12)
def piano(t0,note,vel=0.25,dur=3.0):
    n=int(dur*SR); t=np.arange(n)/SR; f=hz(note)
    w=sum((1/k**1.6)*np.sin(2*np.pi*f*k*t*(1+0.0004*k)) for k in range(1,7))
    env=np.minimum(t/0.006,1)*np.exp(-t*(1.6+note/60))
    s=int(t0*SR); out[s:s+n]+=vel*w*env
def pad(t0,t1,notes,vel=0.05,att=1.2,rel=1.2):
    n=int((t1-t0)*SR); t=np.arange(n)/SR; w=np.zeros(n)
    for m in notes:
        for d in (-0.003,0,0.003):
            f=hz(m)*(1+d); w+=np.sin(2*np.pi*f*t)+0.3*np.sin(2*np.pi*2*f*t)
    env=np.minimum(t/att,1)*np.minimum((t[-1]-t)/rel,1)
    s=int(t0*SR); out[s:s+n]+=vel*w*env/len(notes)
# A: sorrow, A minor  (1.5 - 14.8)
prog=[(57,[45,52,57,60,64]),(53,[41,48,53,57,60]),(48,[36,48,52,55,60]),(55,[43,50,55,59,62])]
t=1.3; i=0
while t<13.2:
    root,ch=prog[i%4]; pad(t,min(t+3.4,14.0),ch[:3],0.045)
    for j,nt in enumerate([ch[2],ch[3],ch[4],ch[3]]):
        if t+j*0.8<13.4: piano(t+j*0.8,nt,0.16 if j else 0.2)
    t+=3.2; i+=1
# B: tension (14.8 - 27.9): drone + sparse dissonant bell, building
n=int((27.5-14.0)*SR); tt=np.arange(n)/SR
drone=(np.sin(2*np.pi*hz(33)*tt)+0.6*np.sin(2*np.pi*hz(40)*tt)+0.25*np.sin(2*np.pi*hz(46)*tt))
drone*=(0.6+0.4*np.sin(2*np.pi*0.5*tt))*np.minimum(tt/1.5,1)*(0.5+0.5*tt/tt[-1])*np.minimum((tt[-1]-tt)/0.05,1)
out[int(14.0*SR):int(14.0*SR)+n]+=0.06*drone
for k,tb in enumerate(np.arange(14.6,27.2,1.6)):
    piano(tb,[76,77,76,82][k%4],0.07+0.006*k,2.0)
# heartbeat-like low pulse under the consequence card (25.9-27.9)
for tb in np.arange(25.5,27.4,0.55):
    piano(tb,33,0.35,0.5)
# C: warmth, C major (29.6 - 44.8)
prog2=[[48,52,55,60,64,67],[43,50,55,59,62,67],[45,52,57,60,64,69],[41,48,53,57,60,65]]
t=29.0; i=0
while t<41.1:
    ch=prog2[i%4]; pad(t,min(t+3.3,42.7),ch[:3],0.05+0.004*i)
    for j,nt in enumerate([ch[3],ch[4],ch[5],ch[4],ch[3],ch[4]]):
        piano(t+j*0.5,nt+(12 if t>38 and j%2==0 else 0),0.11+0.01*i,2.5)
    t+=3.0; i+=1
pad(42.4,45.0,[48,52,55,60,64],0.09,att=0.5,rel=2.2)
piano(42.5,60,0.2,2.5); piano(42.5,72,0.22,2.5); piano(42.6,76,0.18,2.4); piano(42.7,79,0.16,2.3); piano(43.5,84,0.1,1.5)
out=out[:int(SR*T)]
out=out/np.max(np.abs(out))*0.7
st=np.stack([out,np.roll(out,int(0.012*SR))*0.92],1)
with wave.open(sys.argv[1],'wb') as w:
    w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes((st*32767).astype(np.int16).tobytes())
