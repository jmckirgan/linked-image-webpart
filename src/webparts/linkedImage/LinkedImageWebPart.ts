import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import LinkedImage from './components/LinkedImage';

import {
  CaptionAlignment,
  ILinkedImageProps,
  LinkedImageAlignment
} from './components/ILinkedImageProps';


export interface ILinkedImageWebPartProps {
  imageUrl: string;
  filePickerResult?: IFilePickerResult;
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


export default class LinkedImageWebPart
  extends BaseClientSideWebPart<ILinkedImageWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ILinkedImageProps> =
      React.createElement(LinkedImage, {
        imageUrl: this.properties.imageUrl || '',
        linkUrl: this.properties.linkUrl || '',
        altText: this.properties.altText || '',
        tooltip: this.properties.tooltip || '',
        openInNewTab: this.properties.openInNewTab ?? false,
        imageWidth: this.properties.imageWidth || 600,
        alignment: this.properties.alignment || 'left',
        roundedCorners: this.properties.roundedCorners ?? false,
        enableHoverEffect: this.properties.enableHoverEffect ?? true,
        captionText:
  this.properties.captionText || '',

captionFontSize:
  this.properties.captionFontSize || 18,

captionAlignment:
  this.properties.captionAlignment || 'centre'
      });

    ReactDom.render(
      element,
      this.domElement
    );
  }


  protected onInit(): Promise<void> {
    if (this.properties.openInNewTab === undefined) {
      this.properties.openInNewTab = false;
    }

    if (this.properties.imageWidth === undefined) {
      this.properties.imageWidth = 600;
    }

    if (this.properties.alignment === undefined) {
      this.properties.alignment = 'left';
    }

    if (this.properties.roundedCorners === undefined) {
      this.properties.roundedCorners = false;
    }

    if (this.properties.enableHoverEffect === undefined) {
      this.properties.enableHoverEffect = true;
    }
if (this.properties.captionText === undefined) {
  this.properties.captionText = '';
}

if (this.properties.captionFontSize === undefined) {
  this.properties.captionFontSize = 18;
}

if (this.properties.captionAlignment === undefined) {
  this.properties.captionAlignment = 'centre';
}
    return super.onInit();
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }


  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {
      pages: [
        {
          header: {
            description:
              'Configure the linked image. The link opens directly without the SharePoint image lightbox.'
          },

          groups: [
            {
              groupName: 'Image',

              groupFields: [

                PropertyFieldFilePicker(
                  'filePickerResult',
                  {
                    context: this.context as any,

                    filePickerResult:
                      this.properties.filePickerResult as IFilePickerResult,

                    onPropertyChange:
                      this.onPropertyPaneFieldChanged.bind(this),

                    properties:
                      this.properties,

                    onSave: (
                      result: IFilePickerResult
                    ) => {
                      this.properties.filePickerResult =
                        result;

                      if (result.fileAbsoluteUrl) {
                        this.properties.imageUrl =
                          result.fileAbsoluteUrl;

                        this.render();

                        this.context.propertyPane.refresh();
                      }
                    },

                    onChanged: (
                      result: IFilePickerResult
                    ) => {
                      this.properties.filePickerResult =
                        result;
                    },

                    key:
                      'linkedImageFilePicker',

                    label:
                      'Choose image',

                    buttonLabel:
                      'Browse for image',

                    accepts: [
                      '.jpg',
                      '.jpeg',
                      '.png',
                      '.gif',
                      '.webp',
                      '.svg'
                    ],

                    hideWebSearchTab: true,

                    hideLocalUploadTab: true,

                    includePageLibraries: false
                  }
                ),


                PropertyPaneTextField(
                  'imageUrl',
                  {
                    label:
                      'Or enter an image URL',

                    description:
                      'Optional: paste an image URL manually instead of using Browse.'
                  }
                ),


                PropertyPaneTextField(
                  'altText',
                  {
                    label:
                      'Alternative text',

                    description:
                      'Describe the purpose or content of the image.'
                  }
                ),


                PropertyPaneTextField(
                  'tooltip',
                  {
                    label:
                      'Tooltip',

                    description:
                      'Optional text displayed when a pointer rests over the image.'
                  }
                ),


                PropertyPaneSlider(
                  'imageWidth',
                  {
                    label:
                      'Image width',

                    min: 100,

                    max: 1600,

                    step: 10,

                    showValue: true
                  }
                ),


                PropertyPaneDropdown(
                  'alignment',
                  {
                    label:
                      'Alignment',

                    options: [
                      {
                        key: 'left',
                        text: 'Left'
                      },
                      {
                        key: 'centre',
                        text: 'Centre'
                      },
                      {
                        key: 'right',
                        text: 'Right'
                      }
                    ]
                  }
                ),


                PropertyPaneToggle(
                  'roundedCorners',
                  {
                    label:
                      'Rounded corners',

                    onText:
                      'On',

                    offText:
                      'Off'
                  }
                ),


                PropertyPaneToggle(
                  'enableHoverEffect',
                  {
                    label:
                      'Hover effect',

                    onText:
                      'On',

                    offText:
                      'Off'
                  }
                )
              ]
            },


            {
              groupName: 'Link',

              groupFields: [

                PropertyPaneTextField(
                  'linkUrl',
                  {
                    label:
                      'Destination URL',

                    description:
                      'Enter the page, document or website that should open.'
                  }
                ),
PropertyPaneTextField(
  'captionText',
  {
    label:
      'Caption',

    description:
      'Optional text displayed underneath the image.'
  }
),


PropertyPaneSlider(
  'captionFontSize',
  {
    label:
      'Caption font size',

    min: 12,

    max: 36,

    step: 1,

    showValue: true
  }
),


PropertyPaneDropdown(
  'captionAlignment',
  {
    label:
      'Caption alignment',

    options: [
      {
        key: 'left',
        text: 'Left'
      },
      {
        key: 'centre',
        text: 'Centre'
      },
      {
        key: 'right',
        text: 'Right'
      }
    ]
  }
),

                PropertyPaneToggle(
                  'openInNewTab',
                  {
                    label:
                      'Open in a new tab',

                    onText:
                      'Yes',

                    offText:
                      'No'
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }
}