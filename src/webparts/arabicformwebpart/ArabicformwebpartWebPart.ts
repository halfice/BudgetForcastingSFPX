import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'ArabicformwebpartWebPartStrings';
import Arabicformwebpart from './components/Arabicformwebpart';
import { IArabicformwebpartProps } from './components/IArabicformwebpartProps';

export interface IArabicformwebpartWebPartProps {
  description: string;
  greetings : string;
}

export default class ArabicformwebpartWebPart extends BaseClientSideWebPart<IArabicformwebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IArabicformwebpartProps > = React.createElement(
      Arabicformwebpart,
      {
        description: this.properties.description,
        greetings :  this.properties.greetings,
        spHttpClient: this.context.spHttpClient,
        pageContext: this.context.pageContext,
        siteurl:this.context.pageContext.web.absoluteUrl,
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
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('greetings', {
                    label: strings.greetings
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
