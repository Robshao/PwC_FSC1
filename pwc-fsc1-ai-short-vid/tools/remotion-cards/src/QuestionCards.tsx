import React from 'react';
import {AbsoluteFill, useCurrentFrame, useVideoConfig} from 'remotion';
import {enter, exitOpacity} from './anim';
import {theme} from './theme';

// 透明背景，疊在影片上；side／vAlign 決定卡片位置，compact 讓卡片小一號（避免擋到人臉）
export type QuestionProps = {seconds?: number; items: string[]; side?: 'left' | 'right'; vAlign?: 'top' | 'center' | 'bottom'; compact?: boolean; stagger?: number};

export const QuestionCards: React.FC<QuestionProps> = ({items, side = 'right', vAlign = 'center', compact = false, stagger = 18}) => {
  const k = compact ? 0.8 : 1;
  const justify = vAlign === 'top' ? 'flex-start' : vAlign === 'bottom' ? 'flex-end' : 'center';
  const frame = useCurrentFrame();
  const {fps, durationInFrames} = useVideoConfig();
  return (
    <AbsoluteFill style={{fontFamily: theme.font, justifyContent: justify, alignItems: side === 'right' ? 'flex-end' : 'flex-start', padding: '80px 110px', opacity: exitOpacity(frame, durationInFrames)}}>
      <div style={{display: 'flex', flexDirection: 'column', gap: 28 * k}}>
        {items.map((q, i) => (
          <div key={i} style={{display: 'flex', alignItems: 'center', gap: 24, background: theme.card, borderLeft: `8px solid ${theme.accent}`, borderRadius: 10, padding: `${26 * k}px ${40 * k}px`, minWidth: 640 * k, ...enter(frame, fps, 6 + i * stagger)}}>
            <span style={{color: theme.accent, fontSize: 44 * k, fontWeight: 700}}>{i + 1}</span>
            <span style={{color: theme.ink, fontSize: 48 * k, fontWeight: 700}}>{q}</span>
          </div>
        ))}
      </div>
    </AbsoluteFill>
  );
};
