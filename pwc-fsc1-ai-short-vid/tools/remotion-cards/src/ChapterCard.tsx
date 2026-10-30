import React from 'react';
import {AbsoluteFill, useCurrentFrame, useVideoConfig} from 'remotion';
import {enter, exitOpacity} from './anim';
import {theme} from './theme';

export type ChapterProps = {seconds?: number; title: string; subtitle?: string};

export const ChapterCard: React.FC<ChapterProps> = ({title, subtitle}) => {
  const frame = useCurrentFrame();
  const {fps, durationInFrames} = useVideoConfig();
  return (
    <AbsoluteFill style={{background: theme.bg, fontFamily: theme.font, justifyContent: 'center', alignItems: 'center', opacity: exitOpacity(frame, durationInFrames)}}>
      <div style={{width: 120, height: 6, background: theme.accent, marginBottom: 40, ...enter(frame, fps)}} />
      <div style={{color: theme.ink, fontSize: 84, fontWeight: 700, letterSpacing: 4, textAlign: 'center', ...enter(frame, fps, 4)}}>{title}</div>
      {subtitle ? <div style={{color: theme.muted, fontSize: 40, marginTop: 28, ...enter(frame, fps, 12)}}>{subtitle}</div> : null}
    </AbsoluteFill>
  );
};
