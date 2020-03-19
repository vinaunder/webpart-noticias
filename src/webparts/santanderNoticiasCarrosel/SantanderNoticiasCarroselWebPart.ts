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

import styles from "./SantanderNoticiasCarroselWebPart.module.scss";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";
import * as strings from "SantanderNoticiasCarroselWebPartStrings";
// import { NoticiasDS } from "./../../services/appLists.service";
import {
  PublishingPage,
  NoticiasCarrosel
} from "./../../interfaces/appLists.interface";
import { SPFXutils } from "./../../services/util.service";

export interface ISantanderNoticiasCarroselWebPartProps {
  description: string;
  SiteUrl: string;
  ListName: string;
  QtdItens: number;
  ReadMore: string;
  ReadMoreOn: boolean;
  Layout: string;
}

export default class SantanderNoticiasCarroselWebPart extends BaseClientSideWebPart<
  ISantanderNoticiasCarroselWebPartProps
> {
  public NoticiasCarrosel: NoticiasCarrosel;
  public render(): void {
    if (this.properties.Layout == "2colunas") {
      this.getAllListItemsCarrosel().then((ret: NoticiasCarrosel): void => {
        this.NoticiasCarrosel = ret;
        const carousel: any = document.createElement("snt-carousel");
        carousel.datasource = this.NoticiasCarrosel;
        carousel.readmore = this.properties.ReadMore;
        carousel.readmoreon = this.properties.ReadMoreOn;
        this.domElement.innerHTML = "";
        this.domElement.appendChild(carousel);
      });
    } else {
    }
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
  public _NoticiasContent: NoticiasCarrosel = {
    Carrosel: this._Noticias,
    Box1: this._Noticias[0],
    Box2: this._Noticias[1]
  };

  public async getAllListItemsCarrosel(): Promise<NoticiasCarrosel> {
    try {
      const w = Web(this.properties.SiteUrl);
      let quantidade = this.properties.QtdItens;
      const rowLimit = `<RowLimit>${quantidade}</RowLimit>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      const camlQuery = `<Query><Where><And><Eq><FieldRef Name='SANDestaquePrincipal' /><Value Type='Boolean'>1</Value></Eq><Eq><FieldRef Name='SANAtivo' /><Value Type='Boolean'>1</Value></Eq></And></Where><OrderBy><FieldRef Name='SANOrdem1' Ascending='True' /></OrderBy></Query>`;
      const r = await w.lists.getByTitle("Pages").getItemsByCAMLQuery(
        {
          ViewXml: `<View>${camlQuery}${queryOptions}${rowLimit}</View>`
        },
        "FieldValuesAsText",
        "EncodedAbsUrl"
      );
      const quantidadeCarrosel = quantidade - 2;
      // var itemNoticias: NoticiasCarrosel;
      // itemNoticias.Carrosel = [];
      // itemNoticias.Box1 = [];
      // itemNoticias.Box2 = [];

      let itemNoticiasCarrosel: PublishingPage[] = [];
      let ItemNoticiaBox1: PublishingPage;
      let ItemNoticiaBox2: PublishingPage;

      // look through the returned items.
      let CountBox1 = 0;
      for (var i = 0; i < r.length; i++) {
        let iconUrl = null;

        const matches = /SANDestaquePub:SW\|(.*?)\r\n/gi.exec(
          r[i].FieldValuesAsText.MetaInfo
        );
        if (matches !== null && matches.length > 1) {
          // this wil be the value of the PublishingPageImage field
          iconUrl = new SPFXutils().extractIMGUrl(matches[1], "noticias");
        }
        if (i < quantidadeCarrosel) {
          itemNoticiasCarrosel.push({
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
            SANDestaqueCarrosel: iconUrl + "?RenditionID=5",
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=7",
            link: r[i].EncodedAbsUrl
          });
          if (i == quantidadeCarrosel - 1) {
            CountBox1++;
          }
        }
        if (CountBox1 == 1) {
          ItemNoticiaBox1 = {
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
            SANDestaqueCarrosel: iconUrl + "?RenditionID=5",
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=6",
            link: r[i].EncodedAbsUrl
          };
        }
        if (CountBox1 == 2) {
          ItemNoticiaBox2 = {
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
            SANDestaqueCarrosel: iconUrl + "?RenditionID=5",
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=6",
            link: r[i].EncodedAbsUrl
          };
          CountBox1++;
        }
        if (CountBox1 == 1) CountBox1++;
      }
      this._NoticiasContent.Carrosel = itemNoticiasCarrosel;
      this._NoticiasContent.Box1 = ItemNoticiaBox1;
      this._NoticiasContent.Box2 = ItemNoticiaBox2;
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
                PropertyPaneTextField("ReadMore", {
                  label: strings.ReadMoreLabel
                }),
                PropertyPaneToggle("ReadMoreOn", {
                  label: strings.ReadMoreOnLabel
                }),
                PropertyPaneChoiceGroup("Layout", {
                  label: strings.LayoutLabel,
                  options: [
                    {
                      key: "Full",
                      text: "Full Banner",
                      imageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png",
                      imageSize: {
                        width: 32,
                        height: 32
                      },
                      selectedImageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png"
                    },
                    {
                      key: "2colunas",
                      text: "2 Columns",
                      imageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png",
                      imageSize: {
                        width: 32,
                        height: 32
                      },
                      selectedImageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png"
                    }
                  ]
                }),
                PropertyPaneSlider("QtdItens", {
                  label: strings.QtdItensLabel,
                  min: 5,
                  max: 12,
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
