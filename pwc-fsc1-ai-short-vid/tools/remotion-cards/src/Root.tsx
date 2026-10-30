import React from 'react';
import {Composition} from 'remotion';
import {ChapterCard} from './ChapterCard';
import {QuestionCards} from './QuestionCards';
import {ServiceCard} from './ServiceCard';
import {EndCard} from './EndCard';
import {BoardInk} from './BoardInk';

// 預設內容只是示意；實際內容用 props/*.json 傳入（見 README）
const base = {fps: 30, width: 1920, height: 1080};

// props 裡有 seconds 就用它決定片長
const bySeconds = ({props}: {props: Record<string, unknown>}) =>
  typeof props.seconds === 'number' ? {durationInFrames: Math.round(props.seconds * 30)} : {};

export const Root: React.FC = () => (
  <>
    <Composition id="ChapterCard" component={ChapterCard} durationInFrames={75} {...base} calculateMetadata={bySeconds}
      defaultProps={{title: '〔章節標題〕', subtitle: '〔副標〕'}} />
    <Composition id="QuestionCards" component={QuestionCards} durationInFrames={120} {...base} calculateMetadata={bySeconds}
      defaultProps={{items: ['〔問題一〕', '〔問題二〕', '〔問題三〕'], side: 'right' as const, stagger: 18}} />
    <Composition id="ServiceCard" component={ServiceCard} durationInFrames={240} {...base} calculateMetadata={bySeconds}
      defaultProps={{title: '〔服務卡標題〕', leftHeader: '〔左欄〕', rightHeader: '〔右欄〕', rows: [{need: '〔需求〕', service: '〔服務〕', tag: 'now' as const}]}} />
    <Composition id="EndCard" component={EndCard} durationInFrames={180} {...base} calculateMetadata={bySeconds}
      defaultProps={{headline: '〔主標語〕', cta: '〔CTA〕', signature: '〔署名〕'}} />
    <Composition id="BoardInk" component={BoardInk} durationInFrames={1} {...base}
      defaultProps={{version: 'B' as const, layer: 'all' as const}} />
  </>
);
