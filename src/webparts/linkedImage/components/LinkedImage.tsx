import * as React from 'react';
import styles from './LinkedImage.module.scss';
import { ILinkedImageProps } from './ILinkedImageProps';

const LinkedImage: React.FC<ILinkedImageProps> = ({
  imageUrl,
  linkUrl,
  altText,
  tooltip,
  openInNewTab,
  imageWidth,
  alignment,
  roundedCorners,
  enableHoverEffect
}) => {

  const hasImage: boolean = imageUrl.trim().length > 0;
  const hasLink: boolean = linkUrl.trim().length > 0;

  const containerClassName: string = [
    styles.linkedImage,
    styles[alignment]
  ].join(' ');

  const imageClassName: string = [
    styles.image,
    roundedCorners ? styles.rounded : '',
    enableHoverEffect ? styles.hoverEffect : ''
  ]
    .filter(Boolean)
    .join(' ');

  if (!hasImage) {
    return (
      <div className={styles.placeholder}>
        <div
          className={styles.placeholderIcon}
          aria-hidden="true"
        >
          Image
        </div>

        <div className={styles.placeholderTitle}>
          Linked Image
        </div>

        <div className={styles.placeholderText}>
          Open the web-part properties and enter an image URL.
        </div>
      </div>
    );
  }

  const image = (
    <img
      className={imageClassName}
      src={imageUrl}
      alt={altText || ''}
      title={tooltip || undefined}
      style={{
        width: `${imageWidth}px`,
        maxWidth: '100%'
      }}
    />
  );

  return (
    <div className={containerClassName}>
      {hasLink ? (
        <a
          className={styles.imageLink}
          href={linkUrl}
          target={openInNewTab ? '_blank' : '_self'}
          rel={openInNewTab ? 'noopener noreferrer' : undefined}
          aria-label={altText || tooltip || 'Open link'}
        >
          {image}
        </a>
      ) : (
        image
      )}
    </div>
  );
};

export default LinkedImage;