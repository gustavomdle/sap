import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SapTodasAsTarefasAdmWebPartStrings';
import SapTodasAsTarefasAdm from './components/SapTodasAsTarefasAdm';
import { ISapTodasAsTarefasAdmProps } from './components/ISapTodasAsTarefasAdmProps';

export interface ISapTodasAsTarefasAdmWebPartProps {
  description: string;
}

export default class SapTodasAsTarefasAdmWebPart extends BaseClientSideWebPart<ISapTodasAsTarefasAdmWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISapTodasAsTarefasAdmProps> = React.createElement(
      SapTodasAsTarefasAdm,
      {
        description: this.properties.description,
        context: this.context,
        siteurl: this.context.pageContext.web.absoluteUrl
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
