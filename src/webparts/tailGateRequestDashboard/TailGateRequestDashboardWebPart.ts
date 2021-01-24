import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import * as strings from 'TailGateRequestDashboardWebPartStrings';
import TailGateRequestDashboard from './components/TailGateRequestDashboard';
import { ITailGateRequestDashboardProps } from './components/ITailGateRequestDashboardProps';
import "./styles.scss";
export interface ITailGateRequestDashboardWebPartProps {
  description: string;
  spcontext:WebPartContext;
  siteUrl:string;
}

export default class TailGateRequestDashboardWebPart extends BaseClientSideWebPart <ITailGateRequestDashboardWebPartProps> {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
     
      sp.setup({
        spfxContext: this.context
      });
    });
    
  }
  public render(): void {
    const element: React.ReactElement<ITailGateRequestDashboardProps> = React.createElement(
      TailGateRequestDashboard,
      {
        description: this.properties.description,
        spcontext:this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
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
