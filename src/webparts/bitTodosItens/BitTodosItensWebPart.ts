import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BitTodosItensWebPartStrings';
import BitTodosItens from './components/BitTodosItens';
import { IBitTodosItensProps } from './components/IBitTodosItensProps';


export interface IBitTodosItensWebPartProps {
  description: string;
  statusBIT: string;
  //statusBIT: string;
}


export default class BitTodosItensWebPart extends BaseClientSideWebPart<IBitTodosItensWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBitTodosItensProps> = React.createElement(
      BitTodosItens,
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
            description: "Escolha o Status do filtro"    
          },    
          groups: [    
            {    
              groupName: "",    
              groupFields: [    
                PropertyPaneTextField('statusBIT', {    
                  label: strings.StatusFieldLabel    
                })    
              ]    
            }    
          ]    
        }    
      ]    
    };    
  }  
}
