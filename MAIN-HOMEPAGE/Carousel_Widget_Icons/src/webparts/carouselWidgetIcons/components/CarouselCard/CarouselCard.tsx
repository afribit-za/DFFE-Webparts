import * as React from 'react';
import styles from './CarouselCard.module.scss';
import { ICarouselItem } from '../../models/IModels';

export interface ICarouselCardProps {
  position: 'center' | 'left' | 'right' | 'hidden';
  label: string;
  item: ICarouselItem | null;
  onReadMore: (item: ICarouselItem) => void;
}

/**
 * A single carousel card displaying an item from a SharePoint list.
 * Positioned as center (active), left, or right with layered opacity.
 */
const CarouselCard: React.FC<ICarouselCardProps> = (props) => {
  const { position, label, item, onReadMore } = props;
  const [imageLoaded, setImageLoaded] = React.useState<boolean>(false);

  const positionClass =
    position === 'center'
      ? styles.center
      : position === 'left'
      ? styles.left
      : position === 'right'
      ? styles.right
      : styles.hidden;

  const handleReadMoreClick = React.useCallback((): void => {
    if (item) {
      onReadMore(item);
    }
  }, [item, onReadMore]);

  const handleImageLoad = React.useCallback((): void => {
    setImageLoaded(true);
  }, []);

  // Format date for display
  const formatDate = (dateStr: string): string => {
    if (!dateStr) return '';
    try {
      const date = new Date(dateStr);
      return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    } catch {
      return dateStr;
    }
  };

  // Empty card fallback
  if (!item) {
    return (
      <div className={`${styles.card} ${positionClass}`}>
        <div className={styles.cardInner}>
          <div className={styles.labelBadge}>{label}</div>
          <div className={styles.emptyCard}>
            <p className={styles.emptyText}>No {label.toLowerCase()} available</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={`${styles.card} ${positionClass}`}>
      <div className={styles.cardInner}>
        {/* Category badge */}
        <div className={styles.labelBadge}>{label}</div>

        {/* Image */}
        {item.imageUrl && (
          <div className={styles.imageContainer}>
            {!imageLoaded && <div className={styles.imagePlaceholder} />}
            <img
              src={item.imageUrl}
              alt={item.title}
              className={`${styles.image} ${imageLoaded ? styles.imageVisible : ''}`}
              onLoad={handleImageLoad}
              loading="lazy"
            />
          </div>
        )}

        {/* Content area */}
        <div className={styles.content}>
          <h4 className={styles.title}>{item.title}</h4>
          <p className={styles.description}>{item.content}</p>
          <button
            className={styles.readMore}
            onClick={handleReadMoreClick}
            type="button"
          >
            Read More
          </button>
          {item.dateUploaded && (
            <span className={styles.date}>{formatDate(item.dateUploaded)}</span>
          )}
        </div>

        {/* Active underline indicator (center card only) */}
        {position === 'center' && <div className={styles.activeIndicator} />}
      </div>
    </div>
  );
};

export default CarouselCard;
