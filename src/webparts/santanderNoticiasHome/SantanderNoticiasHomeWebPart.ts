import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  IPropertyPaneDropdownOption,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./SantanderNoticiasHomeWebPart.module.scss";
import * as strings from "SantanderNoticiasHomeWebPartStrings";

import {
  PublishingPage,
  NoticiasHome,
} from "./../../interfaces/appLists.interface";

import { SPFXutils } from "./../../services/util.service";

import {
  PropertyFieldTermPicker,
  IPickerTerms,
} from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";

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
  Layout: string;
  Caml: string;
  multiSelect: string[];
  Area: IPickerTerms;
}

export default class SantanderNoticiasHomeWebPart extends BaseClientSideWebPart<
  ISantanderNoticiasHomeWebPartProps
> {
  public options: Array<IPropertyPaneDropdownOption> = new Array<
    IPropertyPaneDropdownOption
  >();

  protected onPropertyPaneConfigurationStart(): void {
    if (this.properties.SiteUrl) {
      this.getAllNoticias();
    }
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "Area" && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.options = new Array<IPropertyPaneDropdownOption>();
      const w = Web(this.properties.SiteUrl);
      const viewFields = `<ViewFields>
    <FieldRef Name="Id"></FieldRef>
    <FieldRef Name="Title"></FieldRef>
  </ViewFields>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      let camlQuery = "";
      if (this.properties.Area.length > 0) {
        camlQuery =
          `<Query><Where><And><Eq><FieldRef Name="SANAtivo" /><Value Type="Boolean">1</Value></Eq><Contains><FieldRef Name="SANAreas" /><Value Type="TaxonomyFieldTypeMulti">` +
          this.properties.Area[0].name +
          `</Value></Contains></And></Where><OrderBy><FieldRef Name="Title" Ascending="True" /></OrderBy></Query>`;
      } else {
        camlQuery = `<Query><Where><Eq><FieldRef Name="SANAtivo" /><Value Type="Boolean">1</Value></Eq></Where><OrderBy><FieldRef Name="Title" Ascending="True" /></OrderBy></Query>`;
      }

      w.lists
        .getByTitle("Pages")
        .getItemsByCAMLQuery(
          {
            ViewXml: `<View>${camlQuery}${queryOptions}${viewFields}</View>`,
          },
          "FieldValuesAsText",
          "EncodedAbsUrl"
        )
        .then((itens) => {
          itens.forEach((item) => {
            this.options.push({ key: item.Id, text: item.Title });
          });
          this.context.propertyPane.refresh();
        });
    }
  }

  public async getAllNoticias() {
    this.options = new Array<IPropertyPaneDropdownOption>();
    const w = Web(this.properties.SiteUrl);
    const viewFields = `<ViewFields>
    <FieldRef Name="Id"></FieldRef>
    <FieldRef Name="Title"></FieldRef>
  </ViewFields>`;
    const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
    let camlQuery = "";
    if (this.properties.Area.length > 0) {
      camlQuery =
        `<Query><Where><And><Eq><FieldRef Name="SANAtivo" /><Value Type="Boolean">1</Value></Eq><Contains><FieldRef Name="SANAreas" /><Value Type="TaxonomyFieldTypeMulti">` +
        this.properties.Area[0].name +
        `</Value></Contains></And></Where><OrderBy><FieldRef Name="Title" Ascending="True" /></OrderBy></Query>`;
    } else {
      camlQuery = `<Query><Where><Eq><FieldRef Name="SANAtivo" /><Value Type="Boolean">1</Value></Eq></Where><OrderBy><FieldRef Name="Title" Ascending="True" /></OrderBy></Query>`;
    }

    await w.lists
      .getByTitle("Pages")
      .getItemsByCAMLQuery(
        {
          ViewXml: `<View>${camlQuery}${queryOptions}${viewFields}</View>`,
        },
        "FieldValuesAsText",
        "EncodedAbsUrl"
      )
      .then((itens) => {
        itens.forEach((item) => {
          this.options.push({ key: item.Id, text: item.Title });
        });
        this.context.propertyPane.refresh();
      });
  }

  public NoticiasHome: NoticiasHome;
  public render(): void {
    if (this.properties.Layout == "area") {
      this.getAllListItemsArea().then((ret: NoticiasHome): void => {
        if (ret.Box1.length > 0) {
          this.NoticiasHome = ret;
          const home: any = document.createElement("snt-noticias-home");
          home.datasource = this.NoticiasHome;
          home.tipo = "news-content-area";
          this.domElement.innerHTML = "";
          this.domElement.appendChild(home);
        }
      });
    } else {
      this.getAllListItemsLarge().then((ret: NoticiasHome): void => {
        if (ret.Box1.length > 0) {
          this.NoticiasHome = ret;
          const home: any = document.createElement("snt-noticias-home");
          home.datasource = this.NoticiasHome;
          home.tipo = "news-content";
          this.domElement.innerHTML = "";
          this.domElement.appendChild(home);
        }
      });
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
      link: "url",
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
      link: "url",
    },
  ];
  public _NoticiasContent: NoticiasHome = {
    Box1: this._Noticias,
    Box2: this._Noticias[1],
  };

  public async getAllListItemsLarge(): Promise<NoticiasHome> {
    try {
      const w = Web(this.properties.SiteUrl);
      let quantidade = this.properties.QtdItens;
      const rowLimit = `<RowLimit>${quantidade}</RowLimit>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      // const camlQuery = this.properties.Caml;
      var camlQuery = '<Query><Where><In><FieldRef Name="ID" /><Values>';

      this.properties.multiSelect.forEach((item) => {
        camlQuery += '<Value Type="Number">' + item + "</Value>";
      });

      camlQuery += "</Values></In></Where></Query>";
      const r = await w.lists.getByTitle("Pages").getItemsByCAMLQuery(
        {
          ViewXml: `<View>${camlQuery}${queryOptions}${rowLimit}</View>`,
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
            link: r[i].EncodedAbsUrl,
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
            link: r[i].EncodedAbsUrl,
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

  public async getAllListItemsArea(): Promise<NoticiasHome> {
    try {
      const w = Web(this.properties.SiteUrl);
      let quantidade = this.properties.QtdItens;
      const rowLimit = `<RowLimit>${quantidade}</RowLimit>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      // const camlQuery = this.properties.Caml;
      var camlQuery = '<Query><Where><In><FieldRef Name="ID" /><Values>';

      this.properties.multiSelect.forEach((item) => {
        camlQuery += '<Value Type="Number">' + item + "</Value>";
      });

      camlQuery += "</Values></In></Where></Query>";
      const r = await w.lists.getByTitle("Pages").getItemsByCAMLQuery(
        {
          ViewXml: `<View>${camlQuery}${queryOptions}${rowLimit}</View>`,
        },
        "FieldValuesAsText",
        "EncodedAbsUrl"
      );
      var itemBox1: PublishingPage[] = [];
      var itemBox2: PublishingPage;
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
        itemBox1.push({
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
          link: r[i].EncodedAbsUrl,
        });
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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("SiteUrl", {
                  label: strings.SiteUrlLabel,
                }),
                PropertyPaneTextField("ListName", {
                  label: strings.ListNameLabel,
                }),
                PropertyFieldTermPicker("Area", {
                  label: "Selecione a Área",
                  panelTitle: "Selecione a Área",
                  initialValues: this.properties.Area,
                  allowMultipleSelections: true,
                  excludeSystemGroup: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  limitByGroupNameOrID:
                    "Site Collection - santandernet.sharepoint.com-sites-IPPortugal",
                  limitByTermsetNameOrID: "Areas",
                  key: "termSetsPickerFieldId",
                }),
                PropertyFieldMultiSelect("multiSelect", {
                  key: "multiSelect",
                  label: "Destaques",
                  options: this.options,
                  selectedKeys: this.properties.multiSelect,
                }),
              ],
            },
            {
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneSlider("QtdItens", {
                  label: strings.QtdItensLabel,
                  min: 1,
                  max: 3,
                  showValue: true,
                }),
                PropertyPaneChoiceGroup("Layout", {
                  label: strings.LayoutLabel,
                  options: [
                    {
                      key: "area",
                      text: "Noticias",
                      imageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png",
                      imageSize: {
                        width: 32,
                        height: 32,
                      },
                      selectedImageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png",
                    },
                    {
                      key: "large",
                      text: "Noticias Large",
                      imageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png",
                      imageSize: {
                        width: 32,
                        height: 32,
                      },
                      selectedImageSrc:
                        "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/choicegroup-bar-selected.png",
                    },
                  ],
                }),
                PropertyPaneTextField("Caml", {
                  label: strings.CamlLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
