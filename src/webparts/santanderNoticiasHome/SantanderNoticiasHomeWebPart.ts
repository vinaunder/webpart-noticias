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

import styles from "./SantanderNoticiasHomeWebPart.module.scss";
import * as strings from "SantanderNoticiasHomeWebPartStrings";

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

export interface ISantanderNoticiasHomeWebPartProps {
  description: string;
  SiteUrl: string;
  ListName: string;
  QtdItens: number;
}

export default class SantanderNoticiasHomeWebPart extends BaseClientSideWebPart<
  ISantanderNoticiasHomeWebPartProps
> {
  public NoticiasHome: NoticiasHome;
  public render(): void {
    this.getAllListItemsCarrosel().then((ret: NoticiasHome): void => {
      this.NoticiasHome = ret;
      const home: any = document.createElement("snt-noticias-home");
      home.datasource = this.NoticiasHome;
      this.domElement.innerHTML = "";
      this.domElement.appendChild(home);
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
  public _NoticiasContent: NoticiasHome = {
    Box1: this._Noticias,
    Box2: this._Noticias[1]
  };

  public async getAllListItemsCarrosel(): Promise<NoticiasHome> {
    try {
      const w = Web(this.properties.SiteUrl);
      let quantidade = this.properties.QtdItens;
      const rowLimit = `<RowLimit>${quantidade}</RowLimit>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      const camlQuery = `<Query><Where><And><Eq><FieldRef Name='SANDestaqueHome' /><Value Type='Boolean'>1</Value></Eq><Eq><FieldRef Name='SANAtivo' /><Value Type='Boolean'>1</Value></Eq></And></Where><OrderBy><FieldRef Name='SANOrdem1' Ascending='True' /></OrderBy></Query>`;
      const r = await w.lists.getByTitle("Pages").getItemsByCAMLQuery(
        {
          ViewXml: `<View>${camlQuery}${queryOptions}${rowLimit}</View>`
        },
        "FieldValuesAsText",
        "EncodedAbsUrl"
      );
      const quantidadeCarrosel = quantidade - 1;
      var itemBox1: PublishingPage[] = [];
      var itemBox2: PublishingPage;
      let CountBox1 = 0;
      for (var i = 0; i < r.length; i++) {
        let iconUrl = null;
        console.log("i:", i);
        const matches = /SANDestaquePub:SW\|(.*?)\r\n/gi.exec(
          r[i].FieldValuesAsText.MetaInfo
        );
        if (matches !== null && matches.length > 1) {
          // this wil be the value of the PublishingPageImage field
          iconUrl = new SPFXutils().extractIMGUrl(matches[1], "noticias");
        }
        if (i < quantidadeCarrosel) {
          itemBox1.push({
            Id: r[i].ID,
            Title: r[i].Title,
            Created: r[i].Created,
            CreatedBy: null,
            SANSinopse1: r[i].SANSinopse1,
            SANAreas: r[i].SANAreas,
            SANAtivo: r[i].SANAtivo,
            SANCategorias: ["Categoria"], //r[i].SANCategoriasd,
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
          if (i == quantidadeCarrosel - 1) {
            CountBox1++;
          }
        }
        if (CountBox1 == 1) {
          itemBox2 = {
            Id: r[i].ID,
            Title: r[i].Title,
            Created: r[i].Created,
            CreatedBy: null,
            SANSinopse1: r[i].SANSinopse1,
            SANAreas: r[i].SANAreas,
            SANAtivo: r[i].SANAtivo,
            SANCategorias: ["Categoria"], //r[i].SANCategoriasd,
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
          };
        }
      }
      this._NoticiasContent.Box1 = itemBox1;
      this._NoticiasContent.Box2 = itemBox2;
      return this._NoticiasContent;
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
                PropertyPaneSlider("QtdItens", {
                  label: strings.QtdItensLabel,
                  min: 1,
                  max: 3,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
