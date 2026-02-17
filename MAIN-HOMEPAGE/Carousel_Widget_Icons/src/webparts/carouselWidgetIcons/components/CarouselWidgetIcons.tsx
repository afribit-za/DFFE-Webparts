import * as React from 'react';
import styles from './CarouselWidgetIcons.module.scss';
import { ICarouselWidgetIconsProps } from './ICarouselWidgetIconsProps';
import { ICarouselItem, IWeatherData } from '../models/IModels';
import { DataService } from '../services/DataService';
import CarouselCard from './CarouselCard/CarouselCard';
import WeatherCard from './WeatherCard/WeatherCard';
import SocialIcon from './SocialIcon/SocialIcon';
import ExpandModal from './ExpandModal/ExpandModal';

/**
 * Returns a time-sensitive greeting based on current hour.
 */
const getGreeting = (): string => {
  const hour: number = new Date().getHours();
  if (hour < 12) return 'Good morning';
  if (hour < 17) return 'Good afternoon';
  return 'Good evening';
};

/**
 * Extracts first name from display name.
 */
const getFirstName = (displayName: string): string => {
  if (!displayName) return '';
  const parts: string[] = displayName.split(' ');
  return parts[0];
};

/**
 * Main orchestrator component for the Carousel + Weather + Social Icons web part.
 */
const CarouselWidgetIcons: React.FC<ICarouselWidgetIconsProps> = (props) => {
  const { context, carouselCards, cities, socialIcons } = props;

  // --- Signed-in user info ---
  const userDisplayName: string = context.pageContext.user.displayName || 'User';
  const firstName: string = getFirstName(userDisplayName);
  const greeting: string = getGreeting();

  // --- State ---
  const [carouselData, setCarouselData] = React.useState<ICarouselItem[][]>([[], [], []]);
  const [activeCardIndex, setActiveCardIndex] = React.useState<number>(0);
  const [weatherData, setWeatherData] = React.useState<IWeatherData[]>([]);
  const [activeCityIndex, setActiveCityIndex] = React.useState<number>(0);
  const [selectedItem, setSelectedItem] = React.useState<ICarouselItem | null>(null);
  const [isCarouselLoading, setIsCarouselLoading] = React.useState<boolean>(true);
  const [isWeatherLoading, setIsWeatherLoading] = React.useState<boolean>(true);

  const dataServiceRef = React.useRef<DataService>(new DataService(context));

  // --- Build flattened carousel items (2 per list = 6 total) ---
  const flatItems = React.useMemo((): Array<{ item: ICarouselItem; label: string }> => {
    const result: Array<{ item: ICarouselItem; label: string }> = [];
    carouselCards.forEach(function (cardConfig, configIndex) {
      const items = carouselData[configIndex] || [];
      // Take up to 2 items per list
      items.slice(0, 2).forEach(function (item) {
        result.push({ item: item, label: cardConfig.cardLabel });
      });
    });
    return result;
  }, [carouselData, JSON.stringify(carouselCards)]);

  // --- Fetch carousel data ---
  React.useEffect(() => {
    let cancelled = false;

    const fetchCarouselData = async (): Promise<void> => {
      setIsCarouselLoading(true);

      try {
        const results: ICarouselItem[][] = await Promise.all(
          carouselCards.map((card) => dataServiceRef.current.getCarouselItems(card))
        );
        if (!cancelled) {
          setCarouselData(results);
        }
      } catch {
        // Silently handle — data will show as empty
      }

      if (!cancelled) {
        setIsCarouselLoading(false);
      }
    };

    fetchCarouselData().catch(function () { /* noop */ });

    return () => {
      cancelled = true;
    };
  // Stringify the full config so any property change (site, list, or column) triggers re-fetch
  }, [JSON.stringify(carouselCards)]);

  // --- Fetch weather data ---
  React.useEffect(() => {
    let cancelled = false;

    const fetchWeatherData = async (): Promise<void> => {
      if (cities.length === 0) {
        setWeatherData([]);
        setIsWeatherLoading(false);
        return;
      }

      setIsWeatherLoading(true);

      try {
        const results = await Promise.all(
          cities.map((city) => dataServiceRef.current.getWeatherData(city))
        );
        const validResults: IWeatherData[] = results.filter(
          (r): r is IWeatherData => r !== null
        );

        if (!cancelled) {
          setWeatherData(validResults);
        }
      } catch {
        // Silently handle — weather will show empty state
      }

      if (!cancelled) {
        setIsWeatherLoading(false);
      }
    };

    fetchWeatherData().catch(function () { /* noop */ });

    return () => {
      cancelled = true;
    };
  }, [cities.join(',')]);

  // --- Auto-rotate carousel every 7 seconds ---
  React.useEffect(() => {
    if (flatItems.length <= 1) return undefined;

    const interval = setInterval(() => {
      setActiveCardIndex((prev) => (prev + 1) % flatItems.length);
    }, 7000);

    return () => clearInterval(interval);
  }, [flatItems.length]);

  // --- Auto-rotate weather every 8 seconds ---
  React.useEffect(() => {
    if (weatherData.length <= 2) return undefined;

    const interval = setInterval(() => {
      setActiveCityIndex((prev) => (prev + 1) % weatherData.length);
    }, 8000);

    return () => clearInterval(interval);
  }, [weatherData.length]);

  // --- Handlers ---
  const handleReadMore = React.useCallback((item: ICarouselItem): void => {
    setSelectedItem(item);
  }, []);

  const handleCloseModal = React.useCallback((): void => {
    setSelectedItem(null);
  }, []);

  // --- Helper: get card position (center/left/right/hidden) ---
  const getCardPosition = (index: number): 'center' | 'left' | 'right' | 'hidden' => {
    const total = flatItems.length;
    if (total === 0) return 'hidden';
    const diff = ((index - activeCardIndex) % total + total) % total;
    if (diff === 0) return 'center';
    if (diff === 1) return 'right';
    if (diff === total - 1) return 'left';
    return 'hidden';
  };

  // --- Helper: get 2 visible weather cities ---
  const getVisibleCities = (): Array<{ data: IWeatherData; isSecondary: boolean }> => {
    if (weatherData.length === 0) return [];
    if (weatherData.length === 1) {
      return [{ data: weatherData[0], isSecondary: false }];
    }
    if (weatherData.length === 2) {
      return [
        { data: weatherData[0], isSecondary: false },
        { data: weatherData[1], isSecondary: true }
      ];
    }
    return [
      { data: weatherData[activeCityIndex % weatherData.length], isSecondary: false },
      { data: weatherData[(activeCityIndex + 1) % weatherData.length], isSecondary: true }
    ];
  };

  // --- Shimmer loading placeholder ---
  const renderShimmer = (): React.ReactElement => (
    <div className={styles.shimmer}>
      <div className={styles.shimmerBlock} />
      <div className={styles.shimmerLine} />
      <div className={`${styles.shimmerLine} ${styles.shimmerLineShort}`} />
      <div className={`${styles.shimmerLine} ${styles.shimmerLineMed}`} />
    </div>
  );

  // --- Empty state ---
  const renderEmptyState = (message: string, hint?: string): React.ReactElement => (
    <div className={styles.emptyState}>
      <div className={styles.emptyStateIcon}>📋</div>
      <p className={styles.emptyStateMessage}>{message}</p>
      {hint && <p className={styles.emptyStateHint}>{hint}</p>}
    </div>
  );

  // --- Render ---
  return (
    <div className={styles.container}>
      {/* Welcome Banner */}
      <div className={styles.welcomeBanner}>
        <span className={styles.welcomeText}>
          {greeting}, <strong className={styles.userName}>{firstName}</strong> — <span className={styles.welcomeHighlight}>welcome back</span>!
        </span>
      </div>

      {/* Content row */}
      <div className={styles.contentRow}>
        {/* LEFT — Carousel */}
        <section className={styles.carouselSection}>
          <div className={styles.carouselWrapper}>
            {isCarouselLoading ? (
              renderShimmer()
            ) : flatItems.length === 0 ? (
              renderEmptyState(
                'No content available',
                'Configure list data sources in web part properties'
              )
            ) : (
              flatItems.map((entry, index) => (
                <CarouselCard
                  key={entry.label + '-' + index}
                  position={getCardPosition(index)}
                  label={entry.label}
                  item={entry.item}
                  onReadMore={handleReadMore}
                />
              ))
            )}
          </div>
        </section>

        {/* MIDDLE — Weather Widget */}
        <section className={styles.weatherSection}>
          <div className={styles.weatherWrapper}>
            {isWeatherLoading ? (
              renderShimmer()
            ) : weatherData.length === 0 ? (
              renderEmptyState(
                'No weather data',
                'Add cities in web part properties'
              )
            ) : (
              getVisibleCities().map((entry, index) => (
                <WeatherCard
                  key={`${entry.data.cityName}-${index}`}
                  data={entry.data}
                  isSecondary={entry.isSecondary}
                />
              ))
            )}
          </div>
        </section>

        {/* FAR RIGHT — Social Icons */}
        {socialIcons.length > 0 && (
          <aside className={styles.socialSection}>
            {socialIcons.map((icon, index) => (
              <SocialIcon
                key={index}
                imageUrl={icon.imageUrl}
                linkUrl={icon.linkUrl}
                tooltip={icon.tooltip}
              />
            ))}
          </aside>
        )}
      </div>

      {/* Modal overlay */}
      {selectedItem && (
        <ExpandModal item={selectedItem} onClose={handleCloseModal} />
      )}
    </div>
  );
};

export default CarouselWidgetIcons;
