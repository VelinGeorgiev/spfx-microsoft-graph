import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'AadHttpClientWebPartStrings';
import AadHttpClientComponent from './components/AadHttpClientComponent';
import { IAadHttpClientProps } from './components/IAadHttpClientProps';

export interface IAadHttpClientWebPartProps {
  description: string;
}

export default class AadHttpClientWebPart extends BaseClientSideWebPart<IAadHttpClientWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAadHttpClientProps > = React.createElement(
      AadHttpClientComponent,
      {
        aadHttpClientFactory: this.context.aadHttpClientFactory
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
