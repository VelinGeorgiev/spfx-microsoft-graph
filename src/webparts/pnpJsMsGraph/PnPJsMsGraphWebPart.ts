import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { graph } from "@pnp/graph";

import * as strings from 'PnPJsMsGraphWebPartStrings';
import PnPJsMsGraph from './components/PnPJsMsGraph';
import { IPnPJsMsGraphProps } from './components/IPnPJsMsGraphProps';

export interface IPnPJsMsGraphWebPartProps {
  description: string;
}

export default class PnPJsMsGraphWebPart extends BaseClientSideWebPart<IPnPJsMsGraphWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      graph.setup({
        spfxContext: this.context
      });
    });
  }
  
  public render(): void {
    const element: React.ReactElement<IPnPJsMsGraphProps > = React.createElement(
      PnPJsMsGraph,
      {
        description: this.properties.description
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
