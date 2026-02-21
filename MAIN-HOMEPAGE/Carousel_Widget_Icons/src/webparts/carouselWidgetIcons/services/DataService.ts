import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICarouselItem, ICarouselCardConfig, IWeatherData } from '../models/IModels';

/**
 * Service class for fetching SharePoint list data and weather API data.
 */
export class DataService {
  private _context: WebPartContext;

  constructor(context: WebPartContext) {
    this._context = context;
  }

  /**
   * Fetches top 5 items from a SharePoint list based on the card configuration.
   */
  public async getCarouselItems(config: ICarouselCardConfig): Promise<ICarouselItem[]> {
    if (!config.listName) {
      return [];
    }

    // Ensure required columns are set
    const titleCol: string = config.titleColumn || 'Title';
    const contentCol: string = config.contentColumn || 'Body';
    const imageCol: string = config.imageColumn || '';
    const attachmentCol: string = config.attachmentColumn || '';
    const dateCol: string = config.dateColumn || 'Created';

    // Use DFFECentral site URL if none specified
    const siteUrl: string = config.siteUrl
      ? config.siteUrl.replace(/\/+$/, '')
      : 'https://afribitholdings.sharepoint.com/sites/DFFECentral';

    // Build the endpoint WITHOUT $select to avoid 400 errors when column names
    // don't match the list schema. SharePoint returns all columns and we map them.
    const listEndpoint: string =
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.listName)}')/items` +
      `?$top=2`;

    // Prioritized endpoints: try with AttachmentFiles expansion + date ordering,
    // fall back as needed
    const endpoints: string[] = [
      `${listEndpoint}&$expand=AttachmentFiles&$orderby=${encodeURIComponent(dateCol)} desc`,
      `${listEndpoint}&$expand=AttachmentFiles&$orderby=Modified desc`,
      `${listEndpoint}&$orderby=Modified desc`
    ];

    try {
      let response: SPHttpClientResponse | undefined;

      for (let i = 0; i < endpoints.length; i++) {
        response = await this._context.spHttpClient.get(
          endpoints[i],
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }
        );
        if (response.ok) break;
      }

      if (!response || !response.ok) {
        return [];
      }

      const data: { value: any[] } = await response.json();

      return data.value.map((item: any) => {
        // Extract origin for constructing full URLs from relative paths
        const originMatch: RegExpMatchArray | null = siteUrl.match(/^(https?:\/\/[^\/]+)/);
        const origin: string = originMatch ? originMatch[1] : '';

        // --- IMAGE EXTRACTION ---
        // SP stores "Image" column as JSON: {"fileName":"Reserved_ImageAttachment_...","originalImageName":"..."}
        // The actual image file is in the AttachmentFiles array matching that fileName.
        let imageUrl: string = '';
        let imageFileName: string = ''; // track so we can skip it in attachments

        const rawImage: any = imageCol ? item[imageCol] : undefined;
        if (rawImage !== undefined && rawImage !== null) {
          if (typeof rawImage === 'string') {
            try {
              const parsed: any = JSON.parse(rawImage); // eslint-disable-line @typescript-eslint/no-explicit-any
              if (parsed.fileName) {
                // SP Image column — find matching file in AttachmentFiles
                imageFileName = parsed.fileName;
                if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
                  for (let af = 0; af < item.AttachmentFiles.length; af++) {
                    if (item.AttachmentFiles[af].FileName === imageFileName) {
                      const relPath: string = item.AttachmentFiles[af].ServerRelativeUrl || '';
                      imageUrl = relPath ? origin + relPath : '';
                      break;
                    }
                  }
                }
              } else {
                // Other JSON formats (serverUrl + serverRelativeUrl, Url, etc.)
                imageUrl = DataService._extractImageFromObject(parsed, siteUrl);
              }
            } catch {
              // Not JSON — treat as plain URL
              imageUrl = DataService._resolveUrl(rawImage.trim(), siteUrl);
            }
          } else if (typeof rawImage === 'object') {
            imageUrl = DataService._extractImageFromObject(rawImage, siteUrl);
          }
        }

        // --- ATTACHMENT EXTRACTION ---
        // Use the AttachmentFiles that are NOT the image file (skip Reserved_ImageAttachment_ files)
        let attachmentUrl: string = '';

        // First check configured attachment column
        const rawAttachment: any = attachmentCol ? item[attachmentCol] : undefined;
        if (rawAttachment) {
          if (typeof rawAttachment === 'string') {
            try {
              const parsed: any = JSON.parse(rawAttachment); // eslint-disable-line @typescript-eslint/no-explicit-any
              attachmentUrl = parsed.Url || parsed.url || parsed.serverRelativeUrl || rawAttachment;
            } catch {
              attachmentUrl = rawAttachment;
            }
          } else if (typeof rawAttachment === 'object') {
            attachmentUrl = rawAttachment.Url || rawAttachment.url || rawAttachment.serverRelativeUrl || '';
          }
        }

        // Fallback: use non-image AttachmentFiles entries
        if (!attachmentUrl && item.AttachmentFiles && item.AttachmentFiles.length > 0) {
          for (let af = 0; af < item.AttachmentFiles.length; af++) {
            const fn: string = item.AttachmentFiles[af].FileName || '';
            // Skip image attachment files (they start with Reserved_ImageAttachment_)
            if (fn.indexOf('Reserved_ImageAttachment_') !== 0) {
              const relUrl: string = item.AttachmentFiles[af].ServerRelativeUrl || '';
              if (relUrl) {
                attachmentUrl = origin + relUrl;
              }
              break;
            }
          }
        }

        // Handle content - strip HTML tags for display
        let content: string = item[contentCol] || '';
        if (typeof content === 'string') {
          content = content.replace(/<[^>]*>/g, '');
        }

        return {
          id: item.Id,
          title: item[titleCol] || '',
          content: content,
          imageUrl: imageUrl,
          attachmentUrl: attachmentUrl,
          dateUploaded: item[dateCol] || ''
        };
      });
    } catch {
      return [];
    }
  }

  /**
   * Extracts image URL from a parsed object.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private static _extractImageFromObject(obj: any, siteUrl: string): string {
    if (!obj) return '';

    // SP Thumbnail/Image column: { serverUrl, serverRelativeUrl }
    if (obj.serverUrl && obj.serverRelativeUrl) {
      return obj.serverUrl + obj.serverRelativeUrl;
    }

    // Just serverRelativeUrl (relative path)
    if (obj.serverRelativeUrl) {
      return DataService._resolveUrl(obj.serverRelativeUrl, siteUrl);
    }

    // Hyperlink/Picture column: { Url, Description }
    if (obj.Url) {
      return obj.Url;
    }
    if (obj.url) {
      return obj.url;
    }

    return '';
  }

  /**
   * Resolves a potentially relative URL to an absolute one.
   */
  private static _resolveUrl(url: string, siteUrl: string): string {
    if (!url) return '';

    // Already absolute
    if (url.indexOf('http://') === 0 || url.indexOf('https://') === 0) {
      return url;
    }

    // Relative path starting with /
    if (url.charAt(0) === '/') {
      // Extract origin from siteUrl
      const match: RegExpMatchArray | null = siteUrl.match(/^(https?:\/\/[^\/]+)/);
      if (match) {
        return match[1] + url;
      }
    }

    return url;
  }

  /**
   * Maps lowercase city names to their Wikipedia article titles.
   * Used to fetch real city photos via the Wikipedia REST summary API
   * (free, no API key, open CORS — works from SharePoint Online).
   */
  private static readonly WIKI_ARTICLE: { [key: string]: string } = {
    'johannesburg':     'Johannesburg',
    'sandton':          'Sandton',
    'pretoria':         'Pretoria',
    'cape town':        'Cape_Town',
    'capetown':         'Cape_Town',
    'durban':           'Durban',
    'bloemfontein':     'Bloemfontein',
    'port elizabeth':   'Gqeberha',
    'gqeberha':         'Gqeberha',
    'polokwane':        'Polokwane',
    'mbombela':         'Mbombela',
    'nelspruit':        'Mbombela',
    'east london':      'East_London,_Eastern_Cape',
    'kimberley':        'Kimberley,_Northern_Cape',
    'pietermaritzburg': 'Pietermaritzburg'
  };

  /**
   * Fetch a city background image URL from the Wikipedia page-summary API.
   * Returns null on any failure so it never blocks the weather data render.
   */
  private async _fetchWikipediaCityImage(cityName: string, cleanCity: string): Promise<string | null> {
    const key: string = cityName.toLowerCase();
    const fallback: string = cleanCity.toLowerCase();
    const article: string =
      DataService.WIKI_ARTICLE[key] ||
      DataService.WIKI_ARTICLE[fallback] ||
      cityName.replace(/ /g, '_');

    const endpoint: string =
      `https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(article)}`;

    try {
      const resp: HttpClientResponse = await this._context.httpClient.get(
        endpoint,
        HttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json' },
          // Prevents SharePoint tenant URL from leaking in the Referer header
          referrerPolicy: 'no-referrer'
        }
      );
      if (!resp.ok) return null;
      const json: any = await resp.json();

      // thumbnail.source is e.g.
      // https://upload.wikimedia.org/wikipedia/commons/thumb/.../320px-name.jpg
      // Bump to 800 px for a crisp card background.
      let imgUrl: string | null = (json && json.thumbnail && json.thumbnail.source) ? json.thumbnail.source : null;
      if (imgUrl) {
        imgUrl = imgUrl.replace(/\/\d+px-/, '/800px-');
      }
      return imgUrl;
    } catch {
      return null;
    }
  }

  /**
   * Known city coordinates to avoid an extra geocoding API call.
   */
  private static readonly KNOWN_CITIES: { [key: string]: { lat: number; lon: number; name: string } } = {
    'johannesburg': { lat: -26.2044, lon: 28.0456, name: 'Johannesburg' },
    'pretoria': { lat: -25.7479, lon: 28.2293, name: 'Pretoria' },
    'cape town': { lat: -33.9249, lon: 18.4241, name: 'Cape Town' },
    'capetown': { lat: -33.9249, lon: 18.4241, name: 'Cape Town' },
    'durban': { lat: -29.8587, lon: 31.0218, name: 'Durban' },
    'bloemfontein': { lat: -29.0852, lon: 26.1596, name: 'Bloemfontein' },
    'port elizabeth': { lat: -33.9608, lon: 25.6022, name: 'Port Elizabeth' },
    'gqeberha': { lat: -33.9608, lon: 25.6022, name: 'Gqeberha' },
    'polokwane': { lat: -23.9045, lon: 29.4689, name: 'Polokwane' },
    'mbombela': { lat: -25.4753, lon: 30.9694, name: 'Mbombela' },
    'nelspruit': { lat: -25.4753, lon: 30.9694, name: 'Nelspruit' },
    'east london': { lat: -33.0292, lon: 27.8546, name: 'East London' },
    'kimberley': { lat: -28.7282, lon: 24.7499, name: 'Kimberley' },
    'pietermaritzburg': { lat: -29.6006, lon: 30.3794, name: 'Pietermaritzburg' }
  };

  /**
   * Convert WMO weather code to a human-readable description.
   */
  private static _wmoDescription(code: number): string {
    if (code === 0) return 'Clear sky';
    if (code === 1) return 'Mainly clear';
    if (code === 2) return 'Partly cloudy';
    if (code === 3) return 'Overcast';
    if (code === 45 || code === 48) return 'Fog';
    if (code >= 51 && code <= 55) return 'Drizzle';
    if (code === 56 || code === 57) return 'Freezing drizzle';
    if (code >= 61 && code <= 65) return 'Rain';
    if (code === 66 || code === 67) return 'Freezing rain';
    if (code >= 71 && code <= 77) return 'Snow';
    if (code >= 80 && code <= 82) return 'Rain showers';
    if (code === 85 || code === 86) return 'Snow showers';
    if (code === 95) return 'Thunderstorm';
    if (code === 96 || code === 99) return 'Thunderstorm with hail';
    return 'Unknown';
  }

  /**
   * Fetches weather data using Open-Meteo (free, no API key, no security risk).
   * 1. Resolves city name to coordinates (known lookup or geocoding API).
   * 2. Fetches current weather from Open-Meteo forecast API.
   */
  public async getWeatherData(cityName: string): Promise<IWeatherData | null> {
    if (!cityName) {
      return null;
    }

    // Check sessionStorage cache (15-minute TTL)
    const cacheKey: string = `cwi_weather_${cityName.toLowerCase().replace(/\s/g, '_')}`;
    const cached: string | null = sessionStorage.getItem(cacheKey);

    if (cached) {
      try {
        const parsed: { data: IWeatherData; timestamp: number } = JSON.parse(cached);
        const fifteenMinutes: number = 15 * 60 * 1000;
        if (Date.now() - parsed.timestamp < fifteenMinutes) {
          return parsed.data;
        }
      } catch {
        // Invalid cache, continue to fetch
      }
    }

    try {
      // Strip country code suffix (e.g. "Johannesburg,ZA" → "Johannesburg")
      const cleanCity: string = cityName.split(',')[0].trim();
      let lat: number;
      let lon: number;
      let displayName: string = cleanCity;

      // Try known cities first (avoids geocoding call)
      const known = DataService.KNOWN_CITIES[cleanCity.toLowerCase()];
      if (known) {
        lat = known.lat;
        lon = known.lon;
        displayName = known.name;
      } else {
        // Geocode via Open-Meteo (free, no key)
        const geoEndpoint: string =
          `https://geocoding-api.open-meteo.com/v1/search` +
          `?name=${encodeURIComponent(cleanCity)}&count=1&language=en&format=json`;

        const geoResponse: HttpClientResponse = await this._context.httpClient.get(
          geoEndpoint,
          HttpClient.configurations.v1,
          { referrerPolicy: 'no-referrer' }  // Prevents tenant URL leaking in Referer header
        );
        if (!geoResponse.ok) {
          return null;
        }
        const geoData: any = await geoResponse.json();
        if (!geoData.results || geoData.results.length === 0) {
          return null;
        }
        lat = geoData.results[0].latitude;
        lon = geoData.results[0].longitude;
        displayName = geoData.results[0].name || cleanCity;
      }

      // Fetch weather + Wikipedia city image in parallel
      const weatherEndpoint: string =
        `https://api.open-meteo.com/v1/forecast` +
        `?latitude=${lat}&longitude=${lon}` +
        `&current=temperature_2m,relative_humidity_2m,weather_code,wind_speed_10m,is_day` +
        `&daily=temperature_2m_max,temperature_2m_min&timezone=auto&forecast_days=1`;

      const [weatherResponse, bgImageUrl]: [HttpClientResponse, string | null] = await Promise.all([
        this._context.httpClient.get(
          weatherEndpoint,
          HttpClient.configurations.v1,
          { referrerPolicy: 'no-referrer' }  // Prevents tenant URL leaking in Referer header
        ),
        this._fetchWikipediaCityImage(displayName, cleanCity)
      ]);

      if (!weatherResponse.ok) {
        return null;
      }

      const weatherJson: any = await weatherResponse.json();
      const current: any = weatherJson.current;
      const daily: any = weatherJson.daily;

      const weatherData: IWeatherData = {
        cityName: displayName,
        temperature: Math.round(current.temperature_2m),
        description: DataService._wmoDescription(current.weather_code),
        weatherCode: current.weather_code,
        humidity: current.relative_humidity_2m,
        windSpeed: Math.round(current.wind_speed_10m * 10) / 10,
        isDay: current.is_day === 1,
        tempHigh: daily && daily.temperature_2m_max && daily.temperature_2m_max[0] !== undefined
          ? Math.round(daily.temperature_2m_max[0]) : undefined,
        tempLow: daily && daily.temperature_2m_min && daily.temperature_2m_min[0] !== undefined
          ? Math.round(daily.temperature_2m_min[0]) : undefined,
        backgroundImageUrl: bgImageUrl || undefined
      };

      // Cache the result
      try {
        sessionStorage.setItem(
          cacheKey,
          JSON.stringify({ data: weatherData, timestamp: Date.now() })
        );
      } catch {
        // SessionStorage full or unavailable
      }

      return weatherData;
    } catch {
      return null;
    }
  }
}
