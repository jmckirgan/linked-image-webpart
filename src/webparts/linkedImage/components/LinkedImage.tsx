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
  enableHoverEffect,
  captionText,
  captionFontSize,
  captionAlignment
}) => {

  const hasImage: boolean =
    imageUrl.trim().length > 0;

  const hasLink: boolean =
    linkUrl.trim().length > 0;

  const hasCaption: boolean =
    captionText.trim().length > 0;

  const containerClassName: string = [
    styles.linkedImage,
    alignment === 'left'
      ? styles.left
      : alignment === 'right'
        ? styles.right
        : styles.centre
  ]
    .filter(Boolean)
    .join(' ');

  const cardClassName: string = [
    styles.card,
    roundedCorners
      ? styles.roundedCard
      : '',
    enableHoverEffect
      ? styles.hoverEffect
      : ''
  ]
    .filter(Boolean)
    .join(' ');

  const captionAlignmentClass: string =
    captionAlignment === 'left'
      ? styles.captionLeft
      : captionAlignment === 'right'
        ? styles.captionRight
        : styles.captionCentre;

  const captionClassName: string = [
    styles.caption,
    captionAlignmentClass
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
          Open the web-part properties and choose an image.
        </div>

      </div>
    );
  }

  const cardContent = (
    <div
      className={cardClassName}
      style={{
        width: `${imageWidth}px`,
        maxWidth: '100%'
      }}
    >

      <img
        className={styles.image}
        src={imageUrl}
        alt={altText || ''}
        title={tooltip || undefined}
      />

      {hasCaption && (
        <div
          className={captionClassName}
          style={{
            fontSize: `${captionFontSize}px`
          }}
        >
          {captionText}
        </div>
      )}

    </div>
  );

  return (
    <div className={containerClassName}>

      {hasLink ? (
        <a
          className={styles.imageLink}
          href={linkUrl}
          target={
            openInNewTab
              ? '_blank'
              : '_self'
          }
          rel={
            openInNewTab
              ? 'noopener noreferrer'
              : undefined
          }
          aria-label={
            altText ||
            captionText ||
            tooltip ||
            'Open link'
          }
        >
          {cardContent}
        </a>
      ) : (
        cardContent
      )}

    </div>
  );
};

export default LinkedImage;