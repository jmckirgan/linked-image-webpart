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

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';


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
        imageWidth: this.properties.imageWidth || 300,
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
      this.properties.imageWidth = 300;
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

private async uploadImageToSiteAssets(
  pickerResult: IFilePickerResult
): Promise<string> {

  const file: File =
    await pickerResult.downloadFileContent();

  if (!file) {
    throw new Error(
      'The selected upload could not be read.'
    );
  }

  const webUrl: string =
    this.context.pageContext.web.absoluteUrl;

  const webServerRelativeUrl: string =
    this.context.pageContext.web.serverRelativeUrl === '/'
      ? ''
      : this.context.pageContext.web.serverRelativeUrl;

  const targetFolderServerRelativeUrl: string =
    `${webServerRelativeUrl}/SiteAssets/Linked Images`;

  /*
   * Ensure the Linked Images folder exists.
   *
   * A failure is ignored here because SharePoint returns
   * an error if the folder already exists.
   */
  const createFolderUrl: string =
    `${webUrl}/_api/web/folders/addUsingPath(decodedurl='` +
    `${this.escapeODataValue(targetFolderServerRelativeUrl)}')`;

  const createFolderResponse: SPHttpClientResponse =
    await this.context.spHttpClient.post(
      createFolderUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept:
            'application/json;odata.metadata=none'
        }
      }
    );

  if (
    !createFolderResponse.ok &&
    createFolderResponse.status !== 409
  ) {
    const folderError: string =
      await createFolderResponse.text();

    /*
     * Do not stop if SharePoint reports that the folder
     * already exists.
     */
    if (
      folderError
        .toLowerCase()
        .indexOf('already exists') === -1
    ) {
      throw new Error(
        `The Linked Images folder could not be created. ` +
        `${folderError}`
      );
    }
  }

  const safeFileName: string =
    this.createSafeFileName(file.name);

  const escapedFolderUrl: string =
    this.escapeODataValue(
      targetFolderServerRelativeUrl
    );

  const escapedFileName: string =
    this.escapeODataValue(safeFileName);

  const uploadUrl: string =
    `${webUrl}/_api/web/` +
    `getfolderbyserverrelativeurl('${escapedFolderUrl}')/` +
    `files/add(overwrite=true,url='${escapedFileName}')`;

  const uploadResponse: SPHttpClientResponse =
    await this.context.spHttpClient.post(
      uploadUrl,
      SPHttpClient.configurations.v1,
      {
        body: file,
        headers: {
          Accept:
            'application/json;odata.metadata=none'
        }
      }
    );

  if (!uploadResponse.ok) {
    const uploadError: string =
      await uploadResponse.text();

    throw new Error(
      `The image could not be uploaded. ${uploadError}`
    );
  }

  /*
   * Return an encoded absolute URL suitable for use
   * as the image src value.
   */
  const encodedFolderPath: string =
    targetFolderServerRelativeUrl
      .split('/')
      .map((part: string): string =>
        encodeURIComponent(part)
      )
      .join('/');

  const encodedFileName: string =
    encodeURIComponent(safeFileName);

  return (
    `${window.location.origin}` +
    `${encodedFolderPath}/` +
    `${encodedFileName}`
  );
}

private escapeODataValue(
  value: string
): string {
  return value.replace(/'/g, "''");
}


private createSafeFileName(
  originalFileName: string
): string {

  const cleanedFileName: string =
    originalFileName
      .trim()
      .replace(/[~"#%&*:<>?/\\{|}]/g, '-')
      .replace(/\s+/g, ' ');

  const lastDotIndex: number =
    cleanedFileName.lastIndexOf('.');

  const baseName: string =
    lastDotIndex > 0
      ? cleanedFileName.substring(
          0,
          lastDotIndex
        )
      : cleanedFileName;

  const extension: string =
    lastDotIndex > 0
      ? cleanedFileName.substring(
          lastDotIndex
        )
      : '';

  /*
   * Add a timestamp so that two uploads with the same
   * name do not accidentally replace one another.
   */
  const timestamp: string =
    new Date()
      .toISOString()
      .replace(/[-:.TZ]/g, '');

  return `${baseName}-${timestamp}${extension}`;
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
                    // PnP FilePicker currently exposes a conflicting SPFx context type.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
                    context: this.context as any,

                    filePickerResult:
                      this.properties.filePickerResult as IFilePickerResult,

                    onPropertyChange:
                      this.onPropertyPaneFieldChanged.bind(this),

                    properties:
                      this.properties,

                    onSave: async (
                      result: IFilePickerResult
                    ): Promise<void> => {
                      this.properties.filePickerResult = result;

                      try {
                        if (result.fileAbsoluteUrl) {
                          this.properties.imageUrl = result.fileAbsoluteUrl;
                        } else {
                          this.properties.imageUrl =
                            await this.uploadImageToSiteAssets(result);
                        }

                        this.render();
                        this.context.propertyPane.refresh();
                      } catch (error: unknown) {
                        const errorMessage: string =
                          error instanceof Error
                            ? error.message
                            : 'An unknown upload error occurred.';

                        console.error(
                          'Linked Image upload failed:',
                          errorMessage
                        );

                        window.alert(
                          `The image could not be uploaded.\n\n` +
                          `${errorMessage}`
                        );
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
hideOneDriveTab: true,
                    hideLocalUploadTab: false,

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