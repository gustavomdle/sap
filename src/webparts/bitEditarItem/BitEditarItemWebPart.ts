import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BitEditarItemWebPartStrings';
import BitEditarItem from './components/BitEditarItem';
import { IBitEditarItemProps } from './components/IBitEditarItemProps';

export interface IBitEditarItemWebPartProps {
  description: string;
}

export default class BitEditarItemWebPart extends BaseClientSideWebPart<IBitEditarItemWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBitEditarItemProps> = React.createElement(
      BitEditarItem,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        addUsersAprovadorEngenharia: [],
        addUsersAprovadorGeral: [],
        addUsersDestinatariosAdicionais: [],
        _addUsersAprovadorEngenharia: []      }
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
