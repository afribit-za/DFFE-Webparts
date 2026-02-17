import * as React from 'react';
import styles from './ExpandModal.module.scss';
import { ICarouselItem } from '../../models/IModels';

export interface IExpandModalProps {
  item: ICarouselItem;
  onClose: () => void;
}

/**
 * Full-content modal popup for carousel items.
 * Displays full image, title, content (no line limit),
 * date, and clickable attachment link.
 */
const ExpandModal: React.FC<IExpandModalProps> = (props) => {
  const { item, onClose } = props;
  const [isVisible, setIsVisible] = React.useState<boolean>(false);

  // Trigger enter animation on mount
  React.useEffect(() => {
    // Small delay to allow CSS transition on mount
    const timer = setTimeout(() => setIsVisible(true), 20);
    return () => clearTimeout(timer);
  }, []);

  // Close with animation
  const handleClose = React.useCallback((): void => {
    setIsVisible(false);
    setTimeout(() => onClose(), 300);
  }, [onClose]);

  // Close on Escape key
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        handleClose();
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [handleClose]);

  // Prevent body scroll when modal is open
  React.useEffect(() => {
    document.body.style.overflow = 'hidden';
    return () => {
      document.body.style.overflow = '';
    };
  }, []);

  // Format date
  const formatDate = (dateStr: string): string => {
    if (!dateStr) return '';
    try {
      const date = new Date(dateStr);
      return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      });
    } catch {
      return dateStr;
    }
  };

  // Get attachment filename from URL
  const getFileName = (url: string): string => {
    if (!url) return 'Download Attachment';
    try {
      const parts = url.split('/');
      return decodeURIComponent(parts[parts.length - 1]) || 'Download Attachment';
    } catch {
      return 'Download Attachment';
    }
  };

  return (
    <div
      className={`${styles.overlay} ${isVisible ? styles.overlayVisible : ''}`}
      onClick={handleClose}
      role="dialog"
      aria-modal="true"
      aria-label={item.title}
    >
      <div
        className={`${styles.modal} ${isVisible ? styles.modalVisible : ''}`}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Close button */}
        <button
          className={styles.closeButton}
          onClick={handleClose}
          aria-label="Close modal"
          type="button"
        >
          ✕
        </button>

        {/* Full image */}
        {item.imageUrl && (
          <div className={styles.imageContainer}>
            <img
              src={item.imageUrl}
              alt={item.title}
              className={styles.image}
            />
          </div>
        )}

        {/* Content body */}
        <div className={styles.body}>
          <h2 className={styles.title}>{item.title}</h2>

          {item.dateUploaded && (
            <span className={styles.date}>{formatDate(item.dateUploaded)}</span>
          )}

          <div className={styles.content}>
            {item.content}
          </div>

          {/* Attachment link */}
          {item.attachmentUrl && (
            <a
              href={item.attachmentUrl}
              target="_blank"
              rel="noopener noreferrer"
              className={styles.attachment}
            >
              <span className={styles.attachmentIcon}>📎</span>
              <span className={styles.attachmentName}>
                {getFileName(item.attachmentUrl)}
              </span>
            </a>
          )}
        </div>
      </div>
    </div>
  );
};

export default ExpandModal;
