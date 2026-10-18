import React from 'react';
import {AbsoluteFill, useCurrentFrame, useVideoConfig} from 'remotion';
import {enter} from './anim';
import {theme} from './theme';

export type EndProps = {seconds?: number; headline: string; cta: string; signature: string; footnote?: string};

export const EndCard: React.FC<EndProps> = ({headline, cta, signature, footnote}) => {
  const frame = useCurrentFrame();
  const {fps} = useVideoConfig();
  return (
    <AbsoluteFill style={{background: '#16181d', fontFamily: theme.font, justifyContent: 'center', alignItems: 'center', textAlign: 'center'}}>
      <div style={{color: theme.ink, fontSize: 72, fontWeight: 700, ...enter(frame, fps)}}>{headline}</div>
      <div style={{background: theme.accent, color: '#fff', fontSize: 46, fontWeight: 700, padding: '18px 44px', marginTop: 50, borderRadius: 6, ...enter(frame, fps, 14)}}>{cta}</div>
      <div style={{color: theme.ink, fontSize: 36, marginTop: 50, ...enter(frame, fps, 26)}}>{signature}</div>
      {footnote ? <div style={{color: '#b9b4ac', fontSize: 24, marginTop: 70, ...enter(frame, fps, 34)}}>{footnote}</div> : null}
    </AbsoluteFill>
  );
};
