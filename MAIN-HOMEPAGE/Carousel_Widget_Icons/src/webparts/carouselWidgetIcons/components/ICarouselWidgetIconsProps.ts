import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICarouselCardConfig, ISocialIconConfig } from '../models/IModels';

/**
 * Props for the main CarouselWidgetIcons component.
 */
export interface ICarouselWidgetIconsProps {
  context: WebPartContext;
  carouselCards: ICarouselCardConfig[];
  cities: string[];
  socialIcons: ISocialIconConfig[];
}
