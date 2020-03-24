import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./SantanderNoticiasCarouselInternoWebPart.module.scss";
import * as strings from "SantanderNoticiasCarouselInternoWebPartStrings";

export interface ISantanderNoticiasCarouselInternoWebPartProps {
  description: string;
  SiteUrl: string;
  ListName: string;
  QtdItens: number;
  Filtro: string;
  Filtroon: boolean;
  Caml: string;
  Titulo: string;
}

export default class SantanderNoticiasCarouselInternoWebPart extends BaseClientSideWebPart<
  ISantanderNoticiasCarouselInternoWebPartProps
> {
  public render(): void {}

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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
