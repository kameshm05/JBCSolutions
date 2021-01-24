import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ReportSignOffWebPartStrings';
import ReportSignOff from './components/ReportSignOff';
import { IReportSignOffProps } from './components/IReportSignOffProps';
import {  WebPartContext } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
export interface IReportSignOffWebPartProps {
  description: string;
  spcontext:WebPartContext;
}

export default class ReportSignOffWebPart extends BaseClientSideWebPart<IReportSignOffWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {     
      sp.setup({
        spfxContext: this.context
      });
    });
    
  }
  public render(): void {
    const element: React.ReactElement<IReportSignOffProps > = React.createElement(
      ReportSignOff,
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
