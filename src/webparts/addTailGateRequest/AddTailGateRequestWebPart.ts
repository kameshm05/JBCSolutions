import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import * as strings from 'AddTailGateRequestWebPartStrings';
import AddTailGateRequest  from './components/AddTailGateRequest'
import { IAddTailGateRequestProps } from './components/IAddTailGateRequestProps';

export interface IAddTailGateRequestWebPartProps {
  description: string;
  spcontext:WebPartContext;
}

export default class AddTailGateRequestWebPart extends BaseClientSideWebPart <IAddTailGateRequestWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
     
      sp.setup({
        spfxContext: this.context
      });
    });
    
  }

  public render(): void {
    const element: React.ReactElement<IAddTailGateRequestProps> = React.createElement(
      AddTailGateRequest,
      {
        description: this.properties.description,
        spcontext:this.context,
        siteURL:this.context.pageContext.web.absoluteUrl
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
