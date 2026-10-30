import {interpolate, spring} from 'remotion';

// 淡入＋往上滑入；delay 以影格為單位
export const enter = (frame: number, fps: number, delay = 0) => {
  const s = spring({frame: frame - delay, fps, config: {damping: 200}});
  return {opacity: s, transform: `translateY(${interpolate(s, [0, 1], [24, 0])}px)`};
};

// 片尾淡出
export const exitOpacity = (frame: number, total: number, len = 12) =>
  interpolate(frame, [total - len, total], [1, 0], {extrapolateLeft: 'clamp', extrapolateRight: 'clamp'});
