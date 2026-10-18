import React from 'react';
import {AbsoluteFill, useCurrentFrame, useVideoConfig} from 'remotion';
import {enter, exitOpacity} from './anim';
import {theme} from './theme';

// tag：now＝已經在做的服務；next＝團隊展望（顏色較淡並加標籤）
export type ServiceRow = {need: string; service: string; detail?: string; tag: 'now' | 'next'};
export type ServiceProps = {seconds?: number; title: string; leftHeader: string; rightHeader: string; rows: ServiceRow[]; footnote?: string};

export const ServiceCard: React.FC<ServiceProps> = ({title, leftHeader, rightHeader, rows, footnote}) => {
  const frame = useCurrentFrame();
  const {fps, durationInFrames} = useVideoConfig();
  return (
    <AbsoluteFill style={{background: theme.bg, fontFamily: theme.font, padding: '90px 130px', opacity: exitOpacity(frame, durationInFrames)}}>
      <div style={{color: theme.ink, fontSize: 60, fontWeight: 700, marginBottom: 44, ...enter(frame, fps)}}>{title}</div>
      <div style={{display: 'flex', color: theme.muted, fontSize: 28, letterSpacing: 2, paddingBottom: 14, borderBottom: `2px solid ${theme.rule}`, ...enter(frame, fps, 6)}}>
        <div style={{width: '42%'}}>{leftHeader}</div>
        <div>{rightHeader}</div>
      </div>
      {rows.map((r, i) => {
        const next = r.tag === 'next';
        return (
          <div key={i} style={{display: 'flex', alignItems: 'center', padding: '22px 0', borderBottom: `1px solid ${theme.rule}`, ...enter(frame, fps, 14 + i * 12)}}>
            <div style={{width: '42%', color: theme.ink, fontSize: 34, paddingRight: 30, opacity: next ? 0.75 : 1}}>{r.need}</div>
            <div style={{flex: 1, display: 'flex', alignItems: 'baseline', gap: 18, flexWrap: 'wrap'}}>
              <span style={{background: next ? 'transparent' : theme.teal, color: next ? theme.accent : theme.bg, border: next ? `2px solid ${theme.accent}` : 'none', fontSize: 30, fontWeight: 700, padding: '4px 16px', borderRadius: 6}}>{r.service}</span>
              {r.detail ? <span style={{color: theme.muted, fontSize: 28}}>{r.detail}</span> : null}
            </div>
          </div>
        );
      })}
      {footnote ? <div style={{position: 'absolute', left: 130, bottom: 60, color: theme.muted, fontSize: 24, ...enter(frame, fps, 30)}}>{footnote}</div> : null}
    </AbsoluteFill>
  );
};
