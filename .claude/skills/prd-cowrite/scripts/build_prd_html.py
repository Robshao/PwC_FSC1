#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""需求文档 md 正文源 → HTML（PRD 模板 HTML 变体）。同源纪律：只手写 md，本脚本生成 HTML。
用法：uv run --no-project --offline python <skill>/scripts/build_prd_html.py src.md out_fragment.html [--standalone out_full.html] [--figures figures.json] [--figs-dir 单图目录] [--check-chairs] [...]
      <skill> 代入真实安装路径；默认图库为该 skill 根下 assets/figures.example.json，不依赖 cwd。
      长参数可写进文件后用 @args.txt 传入（一行一个 token）。
md 约定：<!-- meta k: v --> 头部元信息；<!-- audience 表 --> 受众地图；`## 一、标题` 一级章；`### 1.1 标题` 二级节；`#### ` 三级；
`> 结论：` 紧随标题＝折叠外的一句结论；`{{fig:fN}}` 图（取自图库 json）；节内含 `<!-- auth -->` 才渲染正文，否则黄标待授权；
`<!-- fold -->` 直接层/折叠层分界；bullet 后 4 空格缩进块＝该条「详细」折叠；<mark>…</mark> 待拍（空椅子）；!!…!! 朱红条款；```mermaid 块。
页面样式和渲染逻辑沿用既有工具，展示文案通过参数设置。"""
import sys,re,html,json,os,argparse
LEGEND_DEFAULT='<span><i style="border-color:#2563eb"></i>平台动作 · 主链</span><span><i style="border-color:#ea580c"></i>人的动作</span><span><i style="border-color:#94a3b8;border-top-style:dashed"></i>并行 · 可选 · 回路</span><span><i style="border-color:#B03A28"></i>门禁 · 不可协商</span><span><span class="chairtext">黄底</span> 待拍</span>'
ap=argparse.ArgumentParser(description='需求文档 md → HTML 片段（＋可选单文件 html）',fromfile_prefix_chars='@')
ap.add_argument('src_md',help='源 md')
ap.add_argument('out_fragment_html',help='输出：artifact 片段（无 html 壳）')
ap.add_argument('--standalone',default=None,help='另存单文件 html（带 html 壳）的路径；默认不写')
ap.add_argument('--figures',default=None,help='图库 json（含 css 与 figs）；默认＝skill 根 assets/figures.example.json（不依赖 cwd）')
ap.add_argument('--title',default=None,help='<title> 文本；默认＝md 首行 `# ` 标题')
ap.add_argument('--h1',default=None,help='页头 <h1> 文本；默认＝同 --title')
ap.add_argument('--kicker',default='需求文档',help='页头 h1 上方小字；默认「需求文档」')
ap.add_argument('--subtitle',default='折叠态即简版，每条「详细」按需展开。',help='页头 h1 下方一句导读')
ap.add_argument('--legend-html',default=LEGEND_DEFAULT,help='图例条的 HTML；默认＝四色线＋黄底待拍')
ap.add_argument('--chairs-lede',default='全页待拍项的唯一计数处。',help='页尾待拍清单标题下的一句说明')
ap.add_argument('--footer-src-name',default=None,help='页脚显示的源文件名；默认＝src_md 的文件名')
ap.add_argument('--footer-extra',default='',help='页脚「正文源＝…」句之后追加的文字；默认空')
ap.add_argument('--figs-dir',default=None,help='单图目录：构建前把其中每个 <fid>.json 合并进图库 figs[fid]（只在内存合并，不回写图库文件）；默认不合并')
ap.add_argument('--check-chairs',action='store_true',help='构建后做待拍三数对账（源 <mark> 数 / 页顶「待拍 N」 / 页面可见 chairtext 数减图例样例）；不一致以退出码 3 结束，产物仍写出')
A=ap.parse_args()
src,out=A.src_md,A.out_fragment_html
md=open(src,encoding='utf-8').read()
lib=json.load(open(A.figures or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),'assets','figures.example.json'),encoding='utf-8'))
if A.figs_dir:
    import glob
    for p in sorted(glob.glob(os.path.join(A.figs_dir,'*.json'))):
        lib.setdefault('figs',{})[os.path.splitext(os.path.basename(p))[0]]=json.load(open(p,encoding='utf-8'))
chairs=[]; cur=['']
def inline(s):
    s=html.escape(s,quote=False)
    s=re.sub(r'&lt;mark&gt;(.+?)&lt;/mark&gt;',lambda m:(chairs.append((cur[0],re.sub(r'<.*?>','',m.group(1)))) or f'<span class="chairtext">{m.group(1)}</span>'),s)
    s=re.sub(r'!!(.+?)!!',r'<span class="rqt">\1</span>',s)
    s=re.sub(r'\*\*(.+?)\*\*',r'<b>\1</b>',s)
    s=re.sub(r'`([^`]+)`',r'<code>\1</code>',s)
    s=re.sub(r'§([〇一二三四五六七八九十]+)',r'<a class="xr" href="#c\1">§\1</a>',s)
    return s
def render_body(lines):
    out=[]; lst=None; i=0
    def close():
        nonlocal lst
        if lst: out.append(f'</{lst}>'); lst=None
    while i<len(lines):
        l=lines[i]
        mf=re.match(r'\{\{fig:(f\d+)\}\}',l.strip())
        if mf:
            close(); out.append(fig_html(mf.group(1))); i+=1; continue
        if l.startswith('#### '):
            close(); out.append(f'<h4>{inline(re.sub(r"^[〇一二三四五六七八九十]+之[一二三四五六七八九十]+ · ","",l[5:].strip()))}</h4>'); i+=1; continue
        if l.startswith('```mermaid'):
            close(); j=i+1; buf=[]
            while j<len(lines) and not lines[j].startswith('```'): buf.append(lines[j]); j+=1
            out.append('<div class="fig"><pre class="mermaid">'+html.escape('\n'.join(buf),quote=False)+'</pre></div>'); i=j+1; continue
        if l.startswith('|'):
            close(); rows=[]
            while i<len(lines) and lines[i].startswith('|'):
                rows.append([c.strip() for c in lines[i].strip().strip('|').split('|')]); i+=1
            hdr=rows[0]; body=[r for r in rows[1:] if not all(re.fullmatch(r':?-+:?',c or '-') for c in r)]
            t=['<div class="tw"><table><thead><tr>'+''.join(f'<th>{inline(c)}</th>' for c in hdr)+'</tr></thead><tbody>']
            for r in body:
                cls=' class="chair"' if any('<mark>' in c for c in r) else ''
                t.append(f'<tr{cls}>'+''.join(f'<td>{inline(c)}</td>' for c in r)+'</tr>')
            t.append('</tbody></table></div>'); out.append(''.join(t)); continue
        if l.startswith('> 可判定项'):
            close(); out.append(f'<div class="judge"><span class="lbl">可判定项</span><div>{inline(l.split("：",1)[1].strip() if "：" in l else l[2:])}</div></div>'); i+=1; continue
        if l.startswith('> '):
            close(); out.append(f'<blockquote>{inline(l[2:])}</blockquote>'); i+=1; continue
        m=re.match(r'^[-*] (.+)',l)
        if m:
            if lst!='ul': close(); out.append('<ul>'); lst='ul'
            # 收集紧随其后的缩进块（4 空格）作为该 bullet 的展开
            j=i+1; sub=[]
            while j<len(lines) and (lines[j].startswith('    ') or (lines[j].strip()=='' and j+1<len(lines) and lines[j+1].startswith('    '))):
                sub.append(lines[j][4:] if lines[j].startswith('    ') else ''); j+=1
            if any(x.strip() for x in sub):
                out.append(f'<li>{inline(m.group(1))}<details class="bfold"><summary>详细</summary><div class="bfoldbody">'+render_body(sub)+'</div></details></li>')
            else:
                out.append(f'<li>{inline(m.group(1))}</li>')
            i=j; continue
        m=re.match(r'^\d+[.、] (.+)',l)
        if m:
            if lst!='ol': close(); out.append('<ol class="num">'); lst='ol'
            out.append(f'<li>{inline(m.group(1))}</li>'); i+=1; continue
        if l.strip()=='' or l.startswith('<!--'): close(); i+=1; continue
        close(); out.append(f'<p>{inline(l.strip())}</p>'); i+=1
    close(); return '\n'.join(out)
def fig_html(fid):
    f=lib['figs'][fid]; n=fid[1:]
    return f'<div class="figwrap" id="{fid}"><div class="figcap"><span class="lbl">图 {n}</span> <b>{f["title"]}</b> — {f["caption"]}</div>{f["body"]}</div>'
# ── 头部
_m=re.search(r'<!-- meta\n(.*?)-->',md,flags=re.S); _a=re.search(r'<!-- audience\n(.*?)-->',md,flags=re.S)
if not _m: sys.exit('⛔ 缺头部元信息块：md 里要有一段 <!-- meta（换行）状态: …（换行）受众: …（换行）权威: …（换行）变更: …（换行）-->（键名中文、冒号后一个空格）；见 example.md 开头')
if not _a: sys.exit('⛔ 缺受众地图块：md 里要有一段 <!-- audience（换行）一张 md 表（换行）-->；见 example.md 开头')
if re.search(r'^\t',md,flags=re.M): sys.stderr.write('⚠ 发现 tab 缩进：折叠只认 4 个空格，tab 开头的行不会折成「详细」\n')
meta=dict(re.findall(r'^(\S+?): (.+)$',_m.group(1),flags=re.M))
aud=_a.group(1).strip().split('\n')
aud_html=render_body(aud)
title=re.match(r'# (.+)',md).group(1)
page_title=A.title if A.title is not None else title
page_h1=A.h1 if A.h1 is not None else page_title
footer_src=A.footer_src_name if A.footer_src_name is not None else os.path.basename(src)
body_md=md[md.index('\n## '):]
# ── 章节
chapters=re.split(r'^## ',body_md,flags=re.M)[1:]
parts=[]; nav=[]
for chap in chapters:
    head,_,rest=chap.partition('\n'); chtitle=head.strip()
    ch,_,cname=chtitle.partition('、')
    if not cname: ch,cname='附','附录'
    glyph={'一':'1','二':'2','三':'3','四':'4','附':'附'}.get(ch,ch)
    cur[0]=ch; cid=f'c{ch}'; nav.append(f'<a href="#{cid}">{glyph} {cname}</a>' if cname!='附录' else f'<a href="#{cid}">附录</a>')
    secs=re.split(r'^### ',rest,flags=re.M)
    chap_lines=secs[0].split('\n'); lede=''
    for l in chap_lines:
        if l.startswith('> 结论：'): lede=inline(l[len('> 结论：'):].strip())
    lede_html=('<div class="lede">'+lede+'</div>') if lede else ''
    parts.append(f'<section class="chap" id="{cid}"><div class="shead"><div class="ch">{glyph}</div><h2>{inline(cname)}</h2>{lede_html}</div>')
    for sec in secs[1:]:
        shead,_,sbody=sec.partition('\n'); stitle=shead.strip()
        sid='s'+re.sub(r'[^0-9A-Za-z]','',stitle.split(' ')[0])
        cur[0]=stitle.split(' ')[0]
        lines=sbody.split('\n'); lede=''; figs=[]; rest_lines=[]; auth=False; note=''
        for l in lines:
            if l.startswith('> 结论：') and not lede: lede=inline(l[len('> 结论：'):].strip()); continue
            m=re.match(r'\{\{fig:(f\d+)\}\}',l.strip())
            if m and not auth: figs.append(m.group(1)); continue
            if '<!-- auth -->' in l: auth=True; continue
            m2=re.match(r'<!-- (阶段 \d[^>]*?) -->',l.strip())
            if m2: note=m2.group(1); continue
            rest_lines.append(l)
        has_text=any(x.strip() and not x.startswith('<!--') for x in rest_lines)
        parts.append(f'<div class="sec" id="{sid}"><h3>{inline(stitle)}</h3>'+(f'<p class="lede">{lede}</p>' if lede else ''))
        for f in figs: parts.append(fig_html(f))
        if auth and has_text:
            # 两层：<!-- fold --> 之前直接展示，之后折叠
            if any(x.strip()=='<!-- fold -->' for x in rest_lines):
                k=[n for n,x in enumerate(rest_lines) if x.strip()=='<!-- fold -->'][0]
                direct,folded=rest_lines[:k],rest_lines[k+1:]
                if any(x.strip() and not x.startswith('<!--') for x in direct):
                    parts.append('<div class="direct">'+render_body(direct)+'</div>')
                if any(x.strip() and not x.startswith('<!--') for x in folded):
                    parts.append('<details class="fold"><summary>展开细节</summary><div class="foldbody">'+render_body(folded)+'</div></details>')
            else:
                parts.append('<div class="direct">'+render_body(rest_lines)+'</div>')
        else:
            parts.append(f'<div class="slot">文字：待授权后填{("（"+note.replace("阶段 3 起草：","阶段 3 起草｜").replace("阶段 3：","阶段 3｜")+"）") if note else ""}</div>')
        parts.append('</div>')
    parts.append('</section>')
chair_rows=''.join(f'<tr class="chair"><td>{html.escape(c)}</td><td>{html.escape(t)}</td></tr>' for c,t in chairs)
css=lib['css']+'''
.meta{display:grid;grid-template-columns:repeat(4,1fr);gap:0 24px;margin-top:22px;padding-top:16px;border-top:1px solid #36302A}
.meta div{font-size:13px;color:#A79E8E;line-height:1.65}.meta b{display:block;color:#E7D9C2;font-weight:600;font-family:var(--mono);font-size:11px;letter-spacing:.06em;margin-bottom:2px}
.aud{margin-top:18px}.aud table{font-size:12.5px;color:#A79E8E;border-collapse:collapse}.aud th{font-family:var(--mono);font-size:11px;color:#7D7466;text-align:left;padding:4px 14px 4px 0;border-bottom:1px solid #36302A;font-weight:500}.aud td{padding:5px 14px 5px 0;border-bottom:1px solid #36302A;color:#D9D0C0}
section.chap{margin-top:56px}.shead{display:grid;grid-template-columns:72px 1fr;gap:0 14px;border-top:2px solid var(--rule-strong);padding-top:14px;align-items:start}
.shead .ch{font-family:var(--song);font-size:52px;line-height:.85;font-weight:600;color:var(--key)}.shead h2{font-size:24px;font-weight:600;line-height:1.3;padding-top:2px}
.shead .lede{grid-column:2;color:var(--ink-2);font-size:14px;margin-top:6px;max-width:64em}
.sec{margin-top:34px;padding-left:86px}.sec h3{font-size:17px;font-weight:650;margin:0 0 6px}p.lede{color:var(--ink-2);font-size:14px;margin:0 0 10px;max-width:64em}
.figwrap{margin:12px 0 6px}.figcap{font-size:13px;color:var(--ink-2);margin:0 0 8px;padding:6px 12px;border-left:3px solid var(--rule-strong);background:#F5F2EB;border-radius:0 6px 6px 0;max-width:68em}.figcap .lbl{margin-right:6px;font-family:var(--mono);font-size:11.5px;color:var(--ink-3)}
details.fold{margin:10px 0 0;border:1px solid var(--rule);border-radius:8px;background:var(--surface)}details.fold summary{cursor:pointer;font-family:var(--mono);font-size:12px;color:var(--ink-3);padding:8px 14px;list-style:none}details.fold summary::before{content:"▸ ";}details.fold[open] summary::before{content:"▾ "}
.direct{font-size:14.5px;margin:6px 0 4px}.direct p{margin:8px 0;max-width:68em}.direct h4{font-size:14.5px;font-weight:650;margin:16px 0 6px}.direct ul,.direct ol{margin:6px 0 6px 1.4em}.direct li{margin:3px 0}
.foldbody{padding:4px 18px 16px;font-size:14.5px}
details.bfold{display:inline-block;margin-left:6px;vertical-align:baseline}details.bfold summary{display:inline-block;cursor:pointer;font-family:var(--mono);font-size:11px;font-weight:600;color:#3A342C;background:#E7DFD0;border:0;border-radius:6px;padding:2px 10px;line-height:1.6;list-style:none;box-shadow:0 1px 0 rgba(0,0,0,.08)}details.bfold summary:hover{background:#D9CFBC}details.bfold summary::-webkit-details-marker{display:none}details.bfold[open] summary{background:#1A1512;color:#F5F1E8}
details.bfold .bfoldbody{display:block;margin:8px 0 6px;padding:10px 14px;border-left:2px solid var(--rule);font-size:13.5px;color:var(--ink-2)}details.bfold .bfoldbody h4{font-size:13.5px;margin:8px 0 4px;color:var(--ink)}details.bfold .bfoldbody p{margin:6px 0}details.bfold .bfoldbody table{font-size:13px}.foldbody h4{font-size:14.5px;font-weight:650;margin:16px 0 6px}
.foldbody p{margin:8px 0;max-width:68em}.foldbody ul,.foldbody ol{margin:6px 0 6px 1.4em}.foldbody li{margin:3px 0}
.foldbody blockquote{border-left:3px solid var(--rule);padding:4px 14px;color:var(--ink-2);margin:10px 0}
.judge{display:grid;grid-template-columns:72px 1fr;gap:12px;margin:18px 0 4px;padding:10px 14px;background:#F5F2EB;border-radius:8px;font-size:13.5px;color:var(--ink-2)}
.tw{overflow-x:auto;margin:12px 0}table{border-collapse:collapse;width:100%;font-size:13.5px}th{text-align:left;font-weight:600;font-size:12px;font-family:var(--mono);color:var(--ink-3);padding:7px 9px;border-bottom:2px solid var(--rule-strong);vertical-align:bottom}td{padding:8px 9px;border-bottom:1px solid var(--rule);vertical-align:top;line-height:1.6}
tr.chair td{background:var(--chair-bg)}tr.chair td:first-child{box-shadow:inset 3px 0 0 var(--chair)}
.rqt{background:rgba(176,58,40,.10);color:var(--ink);font-weight:600;padding:0 3px;border-radius:2px}.chairtext{background:var(--chair-bg);box-shadow:0 0 0 2px var(--chair-bg);border-radius:2px}a.xr{border-bottom-style:dotted}
code{font-family:var(--mono);font-size:.92em;background:#F5F2EB;padding:0 4px;border-radius:4px}
.chairband{background:var(--chair-bg);margin-top:72px;padding:40px 0 48px}.chairband .shead{border-top:0;padding-top:0}.chairband .shead .ch{color:var(--chair-mk)}.chairband tr.chair td{background:transparent}.chairband td,.chairband th{border-bottom-color:var(--chair-rule)}
@media(max-width:900px){.sec{padding-left:0}.meta{grid-template-columns:1fr 1fr}}
'''
page=f'''<title>{page_title}</title>
<style>{css}</style>
<header class="band"><div class="wrap">
<div class="kicker">{A.kicker}</div>
<h1>{page_h1}</h1>
<p>{A.subtitle}</p>
<div class="meta"><div><b>状态</b>{html.escape(meta.get("状态",""))}</div><div><b>受众</b>{html.escape(meta.get("受众",""))}</div><div><b>权威</b>{html.escape(meta.get("权威",""))}</div><div><b>变更</b>{html.escape(meta.get("变更",""))}</div></div>
<div class="aud"><div class="kicker" style="margin-bottom:6px">受众地图</div>{aud_html}</div>
</div></header>
<nav class="topnav"><div class="wrap">{''.join(nav)}<a class="chairlink" href="#chairs" style="color:var(--chair-mk)">待拍 {len(chairs)}</a></div></nav>
<main class="wrap">
<div class="legend">{A.legend_html}</div>
{''.join(parts)}
</main>
<div class="chairband" id="chairs"><div class="wrap"><section><div class="shead"><div class="ch">待拍</div><h2>拍板清单 · 空椅子索引</h2><div class="lede">{A.chairs_lede}</div></div>
<div class="tw"><table><thead><tr><th>在哪</th><th>待定什么</th></tr></thead><tbody>{chair_rows}</tbody></table></div></section></div></div>
<footer><div class="wrap">正文源＝{footer_src}（本页由它生成；评审改动回改源再生成）。{A.footer_extra}</div></footer>
'''
open(out,'w',encoding='utf-8').write(page)
# 独立版（飞书/本地打开用）：加 html 壳
standalone='<!doctype html>\n<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'+page.split('</title>')[0]+'</title></head><body>'+page.split('</title>',1)[1]+'</body></html>'
if A.standalone: open(A.standalone,'w',encoding='utf-8').write(standalone)
print('built',out,'chapters',len(chapters),'chairs',len(chairs),'figs',page.count('class="figwrap"'),'folds',page.count('<details'))
if A.check_chairs:
    a=len(re.findall(r'<mark>',md)); b=int(re.search(r'class="chairlink"[^>]*>待拍 (\d+)<',page).group(1))
    c=page.count('class="chairtext"')-A.legend_html.count('class="chairtext"')
    ok=(a==b==c); print('待拍三数对账: 源 %d / 页顶 %d / 页面可见 %d -> %s'%(a,b,c,'一致' if ok else '不一致'))
    if not ok: sys.exit(3)
