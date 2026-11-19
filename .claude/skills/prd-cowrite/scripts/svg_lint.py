#!/usr/bin/env python3
"""有限粗查内联 SVG：H/V 线穿框、文字压线、文字横向出界、文字互压。
用法：uv run --no-project --offline python <skill>/scripts/svg_lint.py page.html
<skill> 代入真实安装路径。退出码：0＝范围内未发现问题，1＝发现问题，2＝未检出内联 SVG。
不替代真实渲染与人工看屏；解析或读取失败也以非零退出。
"""
import re,sys
h=open(sys.argv[1],encoding='utf-8').read()
def attrs(tag):
    return dict(re.findall(r'([\w-]+)="([^"]*)"',tag))
def inherit(s):
    # 把 <g> 上的 font-size / text-anchor 下推到缺省的 <text>
    def fix(m):
        ga=attrs(m.group(1)); body=m.group(2)
        def tfix(t):
            ta=t.group(1); add=''
            for k in ('font-size','text-anchor'):
                if k in ga and k+'=' not in ta: add+=f' {k}="{ga[k]}"'
            return '<text'+ta+add+'>'
        return '<g'+m.group(1)+'>'+re.sub(r'<text([^>]*)>',tfix,body)+'</g>'
    return re.sub(r'<g([^>]*)>(.*?)</g>',fix,s,flags=re.S)
svg_count=0; issue_count=0
for n,m in enumerate(re.finditer(r'<svg.*?</svg>',h,flags=re.S),1):
    svg_count+=1
    s=inherit(m.group(0)); vb=[float(x) for x in attrs(re.match(r'<svg[^>]*>',s).group(0))['viewBox'].split()]
    rects=[]; 
    for r in re.findall(r'<rect[^>]*/>',s):
        a=attrs(r)
        if a.get('class')=='pill': continue
        rects.append((float(a['x']),float(a['y']),float(a['x'])+float(a['width']),float(a['y'])+float(a['height']),a.get('fill','')))
    segs=[]
    for p in re.findall(r'<path[^>]*/>',s):
        a=attrs(p); d=a.get('d','')
        if not d or 'L' in d: continue
        toks=re.findall(r'([MHVL])\s*([-\d.]+)(?:[ ,]([-\d.]+))?',d); x=y=0
        for c,v1,v2 in toks:
            v1=float(v1)
            if c=='M': x,y=v1,float(v2)
            elif c=='H': segs.append(((x,y),(v1,y),d)); x=v1
            elif c=='V': segs.append(((x,y),(x,v1),d)); y=v1
    texts=[]
    for t in re.finditer(r'<text([^>]*)>(.*?)</text>',s,flags=re.S):
        a=attrs(t.group(1))
        if a.get('class')=='pill': continue
        txt=re.sub(r'<.*?>','',t.group(2)); fs=float(a.get('font-size',12))
        # 简单 g 继承由 inherit 粗略下推；缺省 font-size 为 12，不求解 CSS 或完整 SVG 继承
        cjk=sum(1 for ch in txt if ord(ch)>0x2E80); w=cjk*fs*1.0+(len(txt)-cjk)*fs*0.55
        x=float(a.get('x',0)); y=float(a.get('y',0)); anc=a.get('text-anchor','start')
        x0=x-w if anc=='end' else x-w/2 if anc=='middle' else x
        texts.append((x0,y-fs*0.85,x0+w,y+fs*0.2,txt[:28],fs))
    out=[]
    # 1 线段穿框（不含只碰边）
    for (p,q,d) in segs:
        for (x0,y0,x1,y1,f) in rects:
            if x1-x0>=900: continue  # 大底带
            xa,xb=sorted((p[0],q[0])); ya,yb=sorted((p[1],q[1]))
            if p[0]==q[0]:  # 竖线
                if x0+1<p[0]<x1-1 and ya<y1-1 and yb>y0+1: out.append(f'线穿框 竖 x={p[0]} y{ya}-{yb} 穿 rect({x0},{y0},{x1},{y1}) d="{d[:40]}"')
            else:
                if y0+1<p[1]<y1-1 and xa<x1-1 and xb>x0+1: out.append(f'线穿框 横 y={p[1]} x{xa}-{xb} 穿 rect({x0},{y0},{x1},{y1}) d="{d[:40]}"')
    # 2 文字压线（线穿过文字框内部）
    for (tx0,ty0,tx1,ty1,txt,fs) in texts:
        for (p,q,d) in segs:
            xa,xb=sorted((p[0],q[0])); ya,yb=sorted((p[1],q[1]))
            if p[0]==q[0] and tx0+2<p[0]<tx1-2 and ya<ty1-1 and yb>ty0+1: out.append(f'文压竖线 「{txt}」 x={p[0]}')
            if p[1]==q[1] and ty0+2<p[1]<ty1-2 and xa<tx1-2 and xb>tx0+2: out.append(f'文压横线 「{txt}」 y={p[1]}')
        if tx1>vb[2]+2 or tx0<vb[0]-2: out.append(f'文出界 「{txt}」 x{tx0:.0f}-{tx1:.0f} (viewBox 宽 {vb[2]})')
    # 3 文字互压
    for i in range(len(texts)):
        for j in range(i+1,len(texts)):
            a=texts[i]; b=texts[j]
            if a[0]<b[2]-2 and a[2]>b[0]+2 and a[1]<b[3]-2 and a[3]>b[1]+2: out.append(f'文互压 「{a[4]}」×「{b[4]}」')
    issue_count+=len(out)
    print(f'=== svg{n} {len(rects)} rects {len(segs)} segs {len(texts)} texts → {len(out)} 处')
    for o in out: print('  ',o)
if svg_count==0:
    print('未检出内联 SVG（0 个）：没有可检查的 SVG 对象；不代表图形验收通过。外链 SVG、图片与 Mermaid 源码不在本次扫描范围。')
    sys.exit(2)
print(f'有限扫描：{svg_count} 个内联 SVG，{issue_count} 处问题；仅覆盖已支持的几何规则，仍需真实渲染看屏。')
sys.exit(1 if issue_count else 0)
