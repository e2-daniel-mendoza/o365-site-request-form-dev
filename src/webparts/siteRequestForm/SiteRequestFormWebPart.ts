import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import { Site, sp } from "@pnp/sp/presets/all";

import * as strings from 'SiteRequestFormWebPartStrings';
import SiteRequestForm from './components/SiteRequestForm';
import { ISiteRequestFormProps } from './components/ISiteRequestFormProps';
import ThemeVariantService from './services/ThemeVariantService';

//import SiteService from './services/SiteCreationFlow/SiteService';

export interface ISiteRequestFormWebPartProps {
  description: string;
  flowURL: string;
  context: WebPartContext;
}

export default class SiteRequestFormWebPart extends BaseClientSideWebPart<ISiteRequestFormWebPartProps> {

  public render(): void {

    const element: React.ReactElement<ISiteRequestFormProps> = React.createElement(
      SiteRequestForm,
      {
        context: this.context,
        description: this.properties.description,
        optionsListTitle: "Request Options",
        flowURL: this.properties.flowURL,
        themeVariant: ThemeVariantService.themeVariant
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
                }),
                PropertyPaneTextField('flowURL', {
                  label: strings.FlowURLFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onInit(): Promise<void> {
    ThemeVariantService.Initialize(this);
    return super.onInit().then(async _ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
}
