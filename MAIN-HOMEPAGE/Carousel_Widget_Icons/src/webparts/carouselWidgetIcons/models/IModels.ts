/**
 * Represents a single item from a SharePoint list used in the carousel.
 */
export interface ICarouselItem {
  id: number;
  title: string;
  content: string;
  imageUrl: string;
  attachmentUrl: string;
  dateUploaded: string;
}

/**
 * Configuration for a single carousel card data source.
 */
export interface ICarouselCardConfig {
  siteUrl: string;
  listName: string;
  cardLabel: string;
  titleColumn: string;
  contentColumn: string;
  imageColumn: string;
  attachmentColumn: string;
  dateColumn: string;
}

/**
 * Weather data returned from the Open-Meteo API (free, no API key needed).
 */
export interface IWeatherData {
  cityName: string;
  temperature: number;
  description: string;
  weatherCode: number;
  humidity: number;
  windSpeed: number;
  isDay: boolean;
  tempHigh?: number;
  tempLow?: number;
  backgroundImageUrl?: string;
}

/**
 * Configuration for a social media icon link.
 */
export interface ISocialIconConfig {
  imageUrl: string;
  linkUrl: string;
  tooltip: string;
}
