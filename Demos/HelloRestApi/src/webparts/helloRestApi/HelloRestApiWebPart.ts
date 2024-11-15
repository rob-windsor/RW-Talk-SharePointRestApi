import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HelloRestApiWebPartStrings';
import HelloRestApi from './components/HelloRestApi';
import { IHelloRestApiProps } from './components/IHelloRestApiProps';

export interface IHelloRestApiWebPartProps {
  description: string;
}

export default class HelloRestApiWebPart extends BaseClientSideWebPart<IHelloRestApiWebPartProps> {
  public render(): void {
    console.log(this.context);
    const element: React.ReactElement<IHelloRestApiProps> = React.createElement(
      HelloRestApi,
      {
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        domElement: this.domElement,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        legacyPageContext: this.context.pageContext.legacyPageContext
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
