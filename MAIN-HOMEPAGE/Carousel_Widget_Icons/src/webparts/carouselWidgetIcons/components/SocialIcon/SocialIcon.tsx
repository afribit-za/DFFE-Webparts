import * as React from 'react';
import styles from './SocialIcon.module.scss';

export interface ISocialIconProps {
  imageUrl: string;
  linkUrl: string;
  tooltip: string;
}

/**
 * Circular social media icon with hover animation.
 * Opens link in a new tab.
 */
const SocialIcon: React.FC<ISocialIconProps> = (props) => {
  const { imageUrl, linkUrl, tooltip } = props;

  return (
    <a
      href={linkUrl || '#'}
      target="_blank"
      rel="noopener noreferrer"
      className={styles.iconLink}
      title={tooltip}
      aria-label={tooltip || 'Social link'}
    >
      <div className={styles.iconCircle}>
        <img
          src={imageUrl}
          alt={tooltip || 'Social icon'}
          className={styles.iconImage}
          loading="lazy"
        />
      </div>
    </a>
  );
};

export default SocialIcon;
