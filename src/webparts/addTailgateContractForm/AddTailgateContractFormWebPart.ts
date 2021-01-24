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
import * as strings from 'AddTailgateContractFormWebPartStrings';
import AddTailgateContractForm from './components/AddTailgateContractForm';
import { IAddTailgateContractFormProps } from './components/IAddTailgateContractFormProps';

export interface IAddTailgateContractFormWebPartProps {
  description: string;
  spcontext:WebPartContext;
}

export default class AddTailgateContractFormWebPart extends BaseClientSideWebPart<IAddTailgateContractFormWebPartProps> {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
     
      sp.setup({
        spfxContext: this.context
      });
    });
    
  }
  public render(): void {
    const element: React.ReactElement<IAddTailgateContractFormProps > = React.createElement(
      AddTailgateContractForm,
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
