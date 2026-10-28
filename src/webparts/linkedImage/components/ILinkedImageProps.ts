export type LinkedImageAlignment =
  'left' |
  'centre' |
  'right';

export type CaptionAlignment =
  'left' |
  'centre' |
  'right';

export interface ILinkedImageProps {
  imageUrl: string;
  linkUrl: string;
  altText: string;
  tooltip: string;
  openInNewTab: boolean;
  imageWidth: number;
  alignment: LinkedImageAlignment;
  roundedCorners: boolean;
  enableHoverEffect: boolean;

  captionText: string;
  captionFontSize: number;
  captionAlignment: CaptionAlignment;
}