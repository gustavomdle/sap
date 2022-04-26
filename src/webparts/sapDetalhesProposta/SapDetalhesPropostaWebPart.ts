import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SapDetalhesPropostaWebPartStrings';
import SapDetalhesProposta from './components/SapDetalhesProposta';
import { ISapDetalhesPropostaProps } from './components/ISapDetalhesPropostaProps';

export interface ISapDetalhesPropostaWebPartProps {
  description: string;
}

export default class SapDetalhesPropostaWebPart extends BaseClientSideWebPart<ISapDetalhesPropostaWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISapDetalhesPropostaProps> = React.createElement(
      SapDetalhesProposta,
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
