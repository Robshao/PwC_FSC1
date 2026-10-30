import React from 'react';
import {AbsoluteFill, continueRender, delayRender} from 'remotion';
import 'lxgw-wenkai-tc-webfont/lxgwwenkaitc-bold.css';

// 白板手寫內容（透明背景靜態圖）。座標以鏡頭 6b 的白板為基準（854x480 畫面座標 ×2.25）。
// layer：base＝標題與標籤；i1～i3＝右側清單逐行（後製依台詞時間逐行「寫出」）
export type BoardProps = {version: 'A' | 'B'; layer: 'base' | 'i1' | 'i2' | 'i3' | 'all'};

const K = 2.25;
const BLUE = '#1f3f9e';
const RED = '#c8322b';
const font = '"LXGW WenKai TC", "WenQuanYi Zen Hei", sans-serif';

const T: React.FC<{x: number; y: number; size: number; color?: string; anchor?: 'left' | 'center' | 'right'; children: React.ReactNode}> = ({x, y, size, color = BLUE, anchor = 'left', children}) => (
  <div style={{position: 'absolute', left: x * K, top: y * K, fontFamily: font, fontWeight: 700, fontSize: size * K, color, whiteSpace: 'nowrap', lineHeight: 1,
    transform: anchor === 'center' ? 'translateX(-50%)' : anchor === 'right' ? 'translateX(-100%)' : undefined}}>{children}</div>
);

const Underline: React.FC<{x1: number; x2: number; y: number}> = ({x1, x2, y}) => (
  <svg style={{position: 'absolute', left: 0, top: 0}} width={1920} height={1080}>
    <path d={`M ${x1 * K} ${y * K} Q ${((x1 + x2) / 2) * K} ${(y + 2.2) * K} ${x2 * K} ${(y - 0.8) * K}`} stroke={RED} strokeWidth={2.2 * K} fill="none" strokeLinecap="round" />
  </svg>
);

const CONTENT = {
  B: {title: 'AI 保險落地藍圖', cols: ['系統', '數據', '治理'], rows: [] as string[], head: '落地三件事',
      items: [['①', '系統｜核心接得上'], ['②', '數據｜指標看得到'], ['③', '治理｜模型經得起']]},
  A: {title: 'AI 風險分層矩陣', cols: ['製造', '金融', '醫療'], rows: ['高', '中', '低'], head: '三個步驟',
      items: [['①', '評估 AI 模型風險'], ['②', '使用情境分層'], ['③', '商品・核保有依據']]},
};

export const BoardInk: React.FC<BoardProps> = ({version, layer}) => {
  const c = CONTENT[version];
  const show = (l: string) => layer === 'all' || layer === l;
  // 等手寫字型（只下載用到的字）載入完成再截圖
  const [handle] = React.useState(() => delayRender('font'));
  React.useEffect(() => {
    const text = [c.title, c.head, ...c.cols, ...c.rows, ...c.items.flat()].join('');
    document.fonts.load(`700 40px "LXGW WenKai TC"`, text).then(() => document.fonts.ready).then(() => continueRender(handle));
  }, [c, handle]);
  return (
    <AbsoluteFill>
      {show('base') && (
        <>
          <T x={525} y={50} size={21} anchor="center">{c.title}</T>
          <Underline x1={452} x2={598} y={77} />
          {c.cols.map((s, i) => <T key={s} x={[459, 521, 588][i]} y={258} size={16} anchor="center">{s}</T>)}
          {c.rows.map((s, i) => <T key={s} x={414} y={[109, 162, 214][i]} size={16} anchor="right" color={RED}>{s}</T>)}
          <T x={642} y={60} size={15} color={RED}>{c.head}</T>
        </>
      )}
      {c.items.map(([n, s], i) => show(`i${i + 1}`) && (
        <div key={n}>
          <T x={642} y={[96, 132, 168][i]} size={14} color={RED}>{n}</T>
          <T x={660} y={[96, 132, 168][i]} size={14}>{s}</T>
        </div>
      ))}
    </AbsoluteFill>
  );
};
