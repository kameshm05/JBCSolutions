import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {  WebPartContext } from '@microsoft/sp-webpart-base';

import { sp } from "@pnp/sp";
import * as strings from 'ContractorApproveHseWebPartStrings';
import ContractorApproveHse from './components/ContractorApproveHse';
import { IContractorApproveHseProps } from './components/IContractorApproveHseProps';

export interface IContractorApproveHseWebPartProps {
  description: string;
  spcontext:WebPartContext;
}

export default class ContractorApproveHseWebPart extends BaseClientSideWebPart<IContractorApproveHseWebPartProps> {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
     
      sp.setup({
        spfxContext: this.context
      });
    });
    
  }
  public render(): void {
    const element: React.ReactElement<IContractorApproveHseProps > = React.createElement(
      ContractorApproveHse,
      {
        description: this.properties.description,
        spcontext:this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl
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
