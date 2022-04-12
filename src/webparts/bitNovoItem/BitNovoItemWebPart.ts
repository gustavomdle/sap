import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'BitNovoItemWebPartStrings';
import BitNovoItem from './components/BitNovoItem';
import { IBitNovoItemProps } from './components/IBitNovoItemProps';

export interface IBitNovoItemWebPartProps {
  description: string;
}

export default class BitNovoItemWebPart extends BaseClientSideWebPart<IBitNovoItemWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBitNovoItemProps> = React.createElement(
      BitNovoItem,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        addUsersAprovadorEngenharia: [],
        addUsersAprovadorGeral: [],
        addUsersDestinatariosAdicionais: [],
        _addUsersAprovadorEngenharia: []
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
