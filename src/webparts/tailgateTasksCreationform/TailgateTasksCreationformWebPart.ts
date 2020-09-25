import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { WebPartContext } from '@microsoft/sp-webpart-base'; 
import { sp } from "@pnp/sp";
import * as strings from 'TailgateTasksCreationformWebPartStrings';
import TailgateTasksCreationform from './components/TailgateTasksCreationform';
import { ITailgateTasksCreationformProps } from './components/ITailgateTasksCreationformProps';

export interface ITailgateTasksCreationformWebPartProps {
  description: string;
  spcontext:WebPartContext;
  SiteURL:String;
  //currentGroupID:string;
}

export default class TailgateTasksCreationformWebPart extends BaseClientSideWebPart<ITailgateTasksCreationformWebPartProps> {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
     
      sp.setup({
        spfxContext: this.context
      });
    });
    
  }
  public render(): void {
    const element: React.ReactElement<ITailgateTasksCreationformProps > = React.createElement(
      TailgateTasksCreationform,
      {
        description: this.properties.description,
        context:this.context,
        SiteURL:this.context.pageContext.web.absoluteUrl,
        //currentGroupID:this.context.pageContext.site.group.id._guid
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
