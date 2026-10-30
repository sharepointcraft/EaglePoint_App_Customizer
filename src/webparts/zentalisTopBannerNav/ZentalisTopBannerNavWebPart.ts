import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IZentalisTopBannerNavWebPartProps {
  CSSUrl: string;
}

export default class ZentalisTopBannerNavWebPart extends BaseClientSideWebPart<IZentalisTopBannerNavWebPartProps> {

  public render(): void {
    const cssUrl = this.properties.CSSUrl || '';

    if (cssUrl) {
      SPComponentLoader.loadCss(cssUrl);
    }

    this.domElement.innerHTML = `<div></div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Global CSS Path',
              groupFields: [
                PropertyPaneTextField('CSSUrl', {
                  label: 'CSS Url'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
