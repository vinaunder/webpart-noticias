import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./SantanderNoticiasCardsWebPart.module.scss";
import * as strings from "SantanderNoticiasCardsWebPartStrings";

export interface ISantanderNoticiasCardsWebPartProps {
  description: string;
  SiteUrl: string;
  ListName: string;
  QtdItens: number;
  Filtro: string;
  Filtroon: boolean;
  Caml: string;
  Titulo: string;
}
import {
  PublishingPage,
  NoticiasHome
} from "./../../interfaces/appLists.interface";

import { SPFXutils } from "./../../services/util.service";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";

export default class SantanderNoticiasCardsWebPart extends BaseClientSideWebPart<
  ISantanderNoticiasCardsWebPartProps
> {
  public render(): void {
    this.getAllNoticias().then((ret: PublishingPage[]): void => {
      this._Noticias = ret;
      const cards: any = document.createElement("snt-noticias-cards");
      cards.datasource = this._Noticias;
      cards.qtditens = this.properties.QtdItens;
      cards.filtro = this.properties.Filtro;
      cards.showfiltro = this.properties.Filtroon;
      cards.titulo = this.properties.Titulo;
      this.domElement.innerHTML = "";
      this.domElement.appendChild(cards);
    });
  }

  public _Noticias: PublishingPage[] = [
    {
      Id: 1,
      Title: "Noticia Teste",
      Created: "13/02/2020",
      CreatedBy: "vinicius souza",
      SANSubTitulo1: "Sub Titulo Noticia",
      SANSinopse1: "Sinopse Noticia",
      SANResponsavel: "Vinicius Souza",
      SANOrdem1: 1,
      SANDestaquePub: "foto.jpg",
      SANDestaqueHome: true,
      SANDataVigencia: new Date(),
      SANAtivo: true,
      SANAreas: ["Riscos"],
      SANFullHtml: "Conteudo de Noticias",
      SANCategorias: ["Riscos", "Segurança", "CyberDefesa"],
      SANDestaqueCarrosel: "iconUrl",
      SANDestaqueCarrosel2: "iconUrl",
      link: "url"
    },
    {
      Id: 1,
      Title: "Noticia Teste 2",
      Created: "13/02/2020",
      CreatedBy: "vinicius souza",
      SANSubTitulo1: "Sub Titulo Noticia 2",
      SANSinopse1: "Sinopse Noticia 2",
      SANResponsavel: "Vinicius Souza",
      SANOrdem1: 1,
      SANDestaquePub: "foto.jpg",
      SANDestaqueHome: true,
      SANDataVigencia: new Date(),
      SANAtivo: true,
      SANAreas: ["Riscos"],
      SANFullHtml: "Conteudo de Noticias 2",
      SANCategorias: ["Riscos", "Segurança", "CyberDefesa"],
      SANDestaqueCarrosel: "iconUrl",
      SANDestaqueCarrosel2: "iconUrl",
      link: "url"
    }
  ];

  public async getAllNoticias(): Promise<PublishingPage[]> {
    try {
      console.log(this.properties.Caml);
      const w = Web(this.properties.SiteUrl);
      let quantidade = this.properties.QtdItens;
      const rowLimit = `<RowLimit>${quantidade}</RowLimit>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      const camlQuery = `${this.properties.Caml}`;
      const r = await w.lists.getByTitle("Pages").getItemsByCAMLQuery(
        {
          //ViewXml: `<View>${camlQuery}${queryOptions}${rowLimit}</View>`
          ViewXml: `<View>${camlQuery}${queryOptions}</View>`
        },
        "FieldValuesAsText",
        "EncodedAbsUrl"
      );

      let itemNoticias: PublishingPage[] = [];
      let CountBox1 = 0;
      for (var i = 0; i < r.length; i++) {
        let iconUrl = null;
        const matches = /SANDestaquePub:SW\|(.*?)\r\n/gi.exec(
          r[i].FieldValuesAsText.MetaInfo
        );
        if (matches !== null && matches.length > 1) {
          iconUrl = new SPFXutils().extractIMGUrl(matches[1], "noticias");
        }

        itemNoticias.push({
          Id: r[i].ID,
          Title: r[i].Title,
          Created: r[i].Created,
          CreatedBy: null,
          SANSinopse1: r[i].SANSinopse1,
          SANAreas: r[i].SANAreas,
          SANAtivo: r[i].SANAtivo,
          SANCategorias: r[i].SANCategorias,
          SANDestaqueHome: r[i].SANDestaqueHome,
          SANOrdem1: r[i].SANOrdem1,
          SANResponsavel: null,
          SANSubTitulo1: r[i].SANSubTitulo1,
          SANFullHtml: r[i].SANSinopse1,
          SANDataVigencia: r[i].SANDataVigencia,
          SANDestaquePub: iconUrl,
          SANDestaqueCarrosel: iconUrl + "?RenditionID=7",
          SANDestaqueCarrosel2: iconUrl + "?RenditionID=11",
          link: r[i].EncodedAbsUrl
        });
      }
      this._Noticias = itemNoticias;
      console.log("ret", this._Noticias);
      return this._Noticias;
    } catch (e) {
      console.error(e);
    }
  }

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
                PropertyPaneTextField("SiteUrl", {
                  label: strings.SiteUrlLabel
                }),
                PropertyPaneTextField("ListName", {
                  label: strings.ListNameLabel
                })
              ]
            },
            {
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneTextField("Titulo", {
                  label: strings.Titulo
                }),
                PropertyPaneTextField("Filtro", {
                  label: strings.FiltroLabel
                }),
                PropertyPaneToggle("Filtroon", {
                  label: strings.FiltroOnLabel
                }),
                PropertyPaneSlider("QtdItens", {
                  label: strings.QtdItensLabel,
                  min: 3,
                  max: 12,
                  showValue: true
                })
              ]
            },
            {
              groupName: strings.CamlDescription,
              groupFields: [
                PropertyPaneTextField("Caml", {
                  label: strings.CamlLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
