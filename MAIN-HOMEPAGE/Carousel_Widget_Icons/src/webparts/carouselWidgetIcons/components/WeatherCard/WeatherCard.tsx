import * as React from 'react';
import styles from './WeatherCard.module.scss';
import { IWeatherData } from '../../models/IModels';
import WeatherIcon from './WeatherIcon';

export interface IWeatherCardProps {
  data: IWeatherData;
  onNext?: () => void;
  hasMultiple?: boolean;
}

/**
 * Full-bleed weather card with city-photo background, dark-left overlay,
 * matching the reference screenshot design.
 */
const WeatherCard: React.FC<IWeatherCardProps> = (props) => {
  const { data, onNext, hasMultiple } = props;

  const descUpper: string = data.description.toUpperCase();

  // Circular day-progress fraction
  const hour: number = new Date().getHours();
  const dayFraction: number = data.isDay
    ? Math.min(Math.max((hour - 6) / 12, 0), 1)
    : 0;
  const circleR: number = 26;
  const circumference: number = 2 * Math.PI * circleR;
  const dashOffset: number = circumference * (1 - dayFraction);

  // Inline style for the background photo layer
  const bgStyle: React.CSSProperties = data.backgroundImageUrl
    ? { backgroundImage: `url(${data.backgroundImageUrl})` }
    : {};

  return (
    <div className={styles.card}>
      {/* Layer 1: City photo background */}
      <div className={styles.cityPhotoBg} style={bgStyle} aria-hidden="true" />
      {/* Layer 2: Gradient overlay — dark on left, transparent on right */}
      <div className={styles.photoOverlay} aria-hidden="true" />

      {/* Content */}
      <div className={styles.cardContent}>

        {/* TOP ROW: location + circular ring */}
        <div className={styles.topRow}>
          <div className={styles.locationRow}>
            <svg width="10" height="14" viewBox="0 0 10 14" aria-hidden="true" className={styles.pinIcon}>
              <path
                d="M5 0C2.24 0 0 2.24 0 5c0 3.75 5 9 5 9s5-5.25 5-9c0-2.76-2.24-5-5-5zm0 6.75A1.75 1.75 0 1 1 5 3.25 1.75 1.75 0 0 1 5 6.75z"
                fill="white"
              />
            </svg>
            <span className={styles.cityName}>{data.cityName.toUpperCase()}</span>
          </div>

          {/* Circular day/night progress ring */}
          <div className={styles.ringWrap}>
            <svg width="62" height="62" viewBox="0 0 64 64" aria-hidden="true">
              <circle cx="32" cy="32" r={circleR} fill="none" stroke="rgba(255,255,255,0.15)" strokeWidth="4.5" />
              <circle
                cx="32" cy="32" r={circleR}
                fill="none"
                stroke="#F7941D"
                strokeWidth="4.5"
                strokeDasharray={`${circumference}`}
                strokeDashoffset={`${dashOffset}`}
                strokeLinecap="round"
                transform="rotate(-90 32 32)"
              />
              <text x="32" y="29" textAnchor="middle" fill="white" fontSize="9" fontWeight="800" letterSpacing="0.5">
                {data.isDay ? 'DAY' : 'NIGHT'}
              </text>
              {data.tempHigh !== undefined && data.tempLow !== undefined ? (
                <text x="32" y="40" textAnchor="middle" fill="rgba(255,255,255,0.55)" fontSize="8">
                  {data.tempHigh}°/{data.tempLow}°
                </text>
              ) : null}
            </svg>
          </div>
        </div>

        {/* MIDDLE: large temperature */}
        <div className={styles.middleRow}>
          <span className={styles.temperature}>{data.temperature}°</span>
        </div>

        {/* CONDITION ROW: icon | "PARTLY CLOUDY" */}
        <div className={styles.conditionRow}>
          <WeatherIcon weatherCode={data.weatherCode} isDay={data.isDay} size={24} />
          <span className={styles.conditionSep} aria-hidden="true" />
          <span className={styles.description}>{descUpper}</span>
        </div>

        {/* BOTTOM ROW: wind + stacked H / L */}
        <div className={styles.bottomRow}>
          <div className={styles.windWrap}>
            <svg width="14" height="12" viewBox="0 0 24 20" aria-hidden="true">
              <path
                d="M3 6h10a3 3 0 1 0-3-3M3 10h14a3 3 0 1 1-3 3M3 14h7a2 2 0 1 0-2-2"
                fill="none"
                stroke="rgba(255,255,255,0.65)"
                strokeWidth="2"
                strokeLinecap="round"
              />
            </svg>
            <span className={styles.windValue}>{data.windSpeed} KM/H</span>
          </div>
          {data.tempHigh !== undefined && data.tempLow !== undefined && (
            <div className={styles.hiLo}>
              <span className={styles.hiLoItem}>H: {data.tempHigh}°</span>
              <span className={styles.hiLoItem}>L: {data.tempLow}°</span>
            </div>
          )}
        </div>
      </div>

      {/* Next city navigation arrow */}
      {hasMultiple && (
        <button
          className={styles.nextBtn}
          onClick={onNext}
          type="button"
          aria-label="Next city"
        >
          <svg width="8" height="13" viewBox="0 0 8 13" aria-hidden="true">
            <path
              d="M1 1l6 5.5-6 5.5"
              fill="none"
              stroke="white"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </button>
      )}
    </div>
  );
};

export default WeatherCard;
