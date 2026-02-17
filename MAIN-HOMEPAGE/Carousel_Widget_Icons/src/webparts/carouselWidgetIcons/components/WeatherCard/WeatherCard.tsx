import * as React from 'react';
import styles from './WeatherCard.module.scss';
import { IWeatherData } from '../../models/IModels';
import WeatherIcon from './WeatherIcon';

export interface IWeatherCardProps {
  data: IWeatherData;
  isSecondary: boolean;
}

/**
 * Dark glassmorphism weather card with simulated city-photo background.
 * Two cards overlap diagonally — primary (behind, top-left), secondary (front, bottom-right).
 */
const WeatherCard: React.FC<IWeatherCardProps> = (props) => {
  const { data, isSecondary } = props;

  const capitalizedDescription: string = data.description
    .split(' ')
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');

  return (
    <div className={`${styles.card} ${isSecondary ? styles.secondary : styles.primary}`}>
      <div className={styles.cardContent}>
        {/* Header: city name + label */}
        <div className={styles.cardHeader}>
          <h4 className={styles.cityName}>{data.cityName}</h4>
          {/* <span className={styles.headerLabel}>Current</span> */}
        </div>

        {/* Centre: icon + temperature + description */}
        <div className={styles.mainWeather}>
          <WeatherIcon
            weatherCode={data.weatherCode}
            isDay={data.isDay}
            size={52}
          />
          <div className={styles.tempInfo}>
            <span className={styles.temperature}>{data.temperature}°</span>
            <span className={styles.weatherDesc}>{capitalizedDescription}</span>
          </div>
        </div>

        {/* Bottom: humidity + wind */}
        <div className={styles.cardDetails}>
          <div className={styles.detailItem}>
            <svg width="13" height="13" viewBox="0 0 16 16" aria-hidden="true">
              <path d="M8 1 C8 1 3 7 3 10.5 C3 13 5.2 15 8 15 C10.8 15 13 13 13 10.5 C13 7 8 1 8 1Z" fill="rgba(255,255,255,0.65)" />
            </svg>
            <span className={styles.detailValue}>{data.humidity}%</span>
          </div>
          <div className={styles.detailItem}>
            <svg width="13" height="13" viewBox="0 0 16 16" aria-hidden="true">
              <path d="M1 4 Q4 2 7 4 Q10 6 13 4 M1 8 Q4 6 7 8 Q10 10 13 8 M1 12 Q4 10 7 12 Q10 14 13 12" fill="none" stroke="rgba(255,255,255,0.65)" strokeWidth="1.5" strokeLinecap="round" />
            </svg>
            <span className={styles.detailValue}>{data.windSpeed} km/h</span>
          </div>
        </div>
      </div>
    </div>
  );
};

export default WeatherCard;
