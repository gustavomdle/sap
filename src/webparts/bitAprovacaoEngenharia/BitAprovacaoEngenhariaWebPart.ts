import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BitAprovacaoEngenhariaWebPartStrings';
import BitAprovacaoEngenharia from './components/BitAprovacaoEngenharia';
import { IBitAprovacaoEngenhariaProps } from './components/IBitAprovacaoEngenhariaProps';

export interface IBitAprovacaoEngenhariaWebPartProps {
  description: string;
  statusBIT: string;

}

export default class BitAprovacaoEngenhariaWebPart extends BaseClientSideWebPart<IBitAprovacaoEngenhariaWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBitAprovacaoEngenhariaProps> = React.createElement(
      BitAprovacaoEngenharia,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        statusBIT: this.properties.statusBIT
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
