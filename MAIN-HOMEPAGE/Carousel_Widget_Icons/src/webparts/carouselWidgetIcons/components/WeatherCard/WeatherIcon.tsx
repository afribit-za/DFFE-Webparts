import * as React from 'react';

export interface IWeatherIconProps {
  weatherCode: number;
  isDay: boolean;
  size?: number;
}

/**
 * Modern, minimalist SVG weather icons mapped from WMO weather codes.
 * Clean flat design — no cartoons.
 */
const WeatherIcon: React.FC<IWeatherIconProps> = ({ weatherCode, isDay, size }) => {
  const s: number = size || 48;

  // Map WMO code to icon category
  const getCategory = (code: number): string => {
    if (code === 0) return isDay ? 'clear-day' : 'clear-night';
    if (code <= 2) return isDay ? 'partly-day' : 'partly-night';
    if (code === 3) return 'cloudy';
    if (code <= 48) return 'fog';
    if (code <= 57) return 'drizzle';
    if (code <= 67) return 'rain';
    if (code <= 77) return 'snow';
    if (code <= 82) return 'rain';
    if (code <= 86) return 'snow';
    return 'thunder';
  };

  const cat: string = getCategory(weatherCode);

  // Color palette — white/light for dark glassmorphism background
  const SUN = '#FBBF24';
  const SUN_GLOW = 'rgba(251,191,36,0.25)';
  const MOON_FILL = 'rgba(255,255,255,0.85)';
  const CLOUD_FILL = 'rgba(255,255,255,0.7)';
  const CLOUD_DARK = 'rgba(255,255,255,0.5)';
  const RAIN_CLR = 'rgba(255,255,255,0.8)';
  const SNOW_CLR = 'rgba(255,255,255,0.9)';
  const BOLT_CLR = '#FBBF24';
  const FOG_CLR = 'rgba(255,255,255,0.5)';

  /* ---------- sub-elements ---------- */
  const sunEl = (cx: number, cy: number, r: number): React.ReactElement => (
    <g>
      <circle cx={cx} cy={cy} r={r + 4} fill={SUN_GLOW} opacity="0.25" />
      <circle cx={cx} cy={cy} r={r} fill={SUN} />
    </g>
  );

  const moonEl = (cx: number, cy: number, r: number): React.ReactElement => (
    <g>
      <circle cx={cx} cy={cy} r={r} fill={MOON_FILL} />
      <circle cx={cx + r * 0.38} cy={cy - r * 0.28} r={r * 0.78} fill="rgba(15,23,42,0.85)" />
    </g>
  );

  const cloudEl = (tx: number, ty: number, sc: number, color?: string): React.ReactElement => {
    const c: string = color || CLOUD_FILL;
    return (
      <g transform={`translate(${tx},${ty}) scale(${sc})`}>
        <circle cx="12" cy="10" r="7" fill={c} />
        <circle cx="22" cy="7" r="9" fill={c} />
        <circle cx="32" cy="10" r="6" fill={c} />
        <rect x="7" y="10" width="30" height="8" rx="3" fill={c} />
      </g>
    );
  };

  const rainDrops = (): React.ReactElement => (
    <g>
      <line x1="16" y1="33" x2="14" y2="39" stroke={RAIN_CLR} strokeWidth="2" strokeLinecap="round" />
      <line x1="24" y1="34" x2="22" y2="42" stroke={RAIN_CLR} strokeWidth="2" strokeLinecap="round" />
      <line x1="32" y1="33" x2="30" y2="39" stroke={RAIN_CLR} strokeWidth="2" strokeLinecap="round" />
    </g>
  );

  const drizzleDrops = (): React.ReactElement => (
    <g>
      <circle cx="16" cy="36" r="1.5" fill={RAIN_CLR} opacity="0.7" />
      <circle cx="24" cy="38" r="1.5" fill={RAIN_CLR} opacity="0.7" />
      <circle cx="32" cy="36" r="1.5" fill={RAIN_CLR} opacity="0.7" />
    </g>
  );

  const snowDots = (): React.ReactElement => (
    <g>
      <circle cx="15" cy="36" r="2" fill={SNOW_CLR} />
      <circle cx="24" cy="39" r="2" fill={SNOW_CLR} />
      <circle cx="33" cy="36" r="2" fill={SNOW_CLR} />
    </g>
  );

  const boltEl = (): React.ReactElement => (
    <polygon points="23,29 27,29 24,35 28,35 20,46 23,37 19,37" fill={BOLT_CLR} />
  );

  const fogLines = (): React.ReactElement => (
    <g opacity="0.6">
      <line x1="10" y1="33" x2="38" y2="33" stroke={FOG_CLR} strokeWidth="2.5" strokeLinecap="round" />
      <line x1="14" y1="38" x2="34" y2="38" stroke={FOG_CLR} strokeWidth="2.5" strokeLinecap="round" />
      <line x1="12" y1="43" x2="36" y2="43" stroke={FOG_CLR} strokeWidth="2" strokeLinecap="round" />
    </g>
  );

  /* ---------- compose ---------- */
  const renderIcon = (): React.ReactElement => {
    switch (cat) {
      case 'clear-day':
        return <g>{sunEl(24, 24, 11)}</g>;
      case 'clear-night':
        return <g>{moonEl(24, 22, 12)}</g>;
      case 'partly-day':
        return <g>{sunEl(15, 14, 7)}{cloudEl(5, 16, 1)}</g>;
      case 'partly-night':
        return <g>{moonEl(15, 13, 7)}{cloudEl(5, 16, 1)}</g>;
      case 'cloudy':
        return <g>{cloudEl(6, 6, 0.8, '#94A3B8')}{cloudEl(4, 14, 1)}</g>;
      case 'fog':
        return <g>{cloudEl(5, 6, 0.85)}{fogLines()}</g>;
      case 'drizzle':
        return <g>{cloudEl(5, 8, 0.9)}{drizzleDrops()}</g>;
      case 'rain':
        return <g>{cloudEl(5, 6, 1)}{rainDrops()}</g>;
      case 'snow':
        return <g>{cloudEl(5, 6, 1)}{snowDots()}</g>;
      case 'thunder':
        return <g>{cloudEl(5, 4, 1, CLOUD_DARK)}{boltEl()}</g>;
      default:
        return <g>{cloudEl(5, 12, 1)}</g>;
    }
  };

  return (
    <svg
      width={s}
      height={s}
      viewBox="0 0 48 48"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden="true"
      style={{ display: 'block', flexShrink: 0 }}
    >
      {renderIcon()}
    </svg>
  );
};

export default WeatherIcon;
