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
  NoticiasCarrosel,
  NoticiasCarroselInterno,
} from "./../../interfaces/appLists.interface";
import { SPFXutils } from "./../../services/util.service";

import {
  PropertyFieldTermPicker,
  IPickerTerms,
} from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

import { PropertyFieldOrder } from "@pnp/spfx-property-controls/lib/PropertyFieldOrder";

import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";

export interface ISantanderNoticiasCarroselWebPartProps {
  description: string;
  SiteUrl: string;
  ListName: string;
  QtdItens: number;
  ReadMore: string;
  ReadMoreOn: boolean;
  Layout: string;
  Caml: string;
  multiSelect: string[];
  Area: IPickerTerms;
  Tipos: IPickerTerms;
  isProduto: boolean;
  isDate: boolean;
  orderedItems: Array<any>;
  siteurlbanner: string;
  listnamebanner: string;
  bannerurl: string;
  bannerimg: string;
  bannertarget: string;
  bannerid: string;
}

export default class SantanderNoticiasCarroselWebPart extends BaseClientSideWebPart<
  ISantanderNoticiasCarroselWebPartProps
> {
  public options: Array<IPropertyPaneDropdownOption> = new Array<
    IPropertyPaneDropdownOption
  >();

  protected onPropertyPaneConfigurationStart(): void {
    if (this.properties.SiteUrl) {
      this.getAllNoticias();
    }
  }

  protected onPropertyPaneFieldChanged2(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "Tipos" && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.options = new Array<IPropertyPaneDropdownOption>();
      const w = Web(this.properties.SiteUrl);
      const viewFields = `<ViewFields>
    <FieldRef Name="Id"></FieldRef>
    <FieldRef Name="Title"></FieldRef>
  </ViewFields>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      let camlQuery = "";
      if (this.properties.Tipos.length > 0) {
        camlQuery =
          `<Query><Where><And><Eq><FieldRef Name="SANAtivo" /><Value Type="Boolean">1</Value></Eq><Contains><FieldRef Name="SANTipo" /><Value Type="TaxonomyFieldTypeMulti">` +
          this.properties.Tipos[0].name +
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
          this.render();
        });
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
        if (this.properties.Tipos.length > 0) {
          camlQuery =
            `<Query><Where><And><Contains><FieldRef Name="SANAreas" /><Value Type="TaxonomyFieldTypeMulti">` +
            this.properties.Area[0].name +
            `</Value></Contains><Contains><FieldRef Name="SANTipo" /><Value Type="TaxonomyFieldType">` +
            this.properties.Tipos[0].name +
            `</Value></Contains></And></Where></Query>`;
        } else {
          camlQuery =
            `<Query><Where><And><Eq><FieldRef Name="SANAtivo" /><Value Type="Boolean">1</Value></Eq><Contains><FieldRef Name="SANAreas" /><Value Type="TaxonomyFieldTypeMulti">` +
            this.properties.Area[0].name +
            `</Value></Contains></And></Where><OrderBy><FieldRef Name="Title" Ascending="True" /></OrderBy></Query>`;
        }
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
          this.render();
        });
    }
  }

  public getAllNoticias() {
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
        this.render();
      });
  }

  public NoticiasCarrosel: NoticiasCarrosel;
  public NoticiasCarroselInterno: NoticiasCarroselInterno;
  public bannerimg: string;
  public bannerurl: string;
  public bannertarget: string;
  public render(): void {
    if (this.properties.Layout == "2colunas") {
      this.getAllListItemsCarrosel().then((ret: NoticiasCarrosel): void => {
        this.NoticiasCarrosel = ret;
        const carousel: any = document.createElement("snt-carousel");
        carousel.datasource = this.NoticiasCarrosel;
        carousel.readmore = this.properties.ReadMore;
        carousel.readmoreon = this.properties.ReadMoreOn;
        carousel.isProduto = this.properties.isProduto;
        this.domElement.innerHTML = "";
        this.domElement.appendChild(carousel);
      });
    } else if (this.properties.Layout == "interno") {
      this.getAllListItemsCarroseInterno().then(
        (ret: NoticiasCarroselInterno): void => {
          if (ret.Carrosel.length > 0) {
            this.NoticiasCarroselInterno = ret;
            const carousel: any = document.createElement(
              "snt-carousel-interno"
            );
            carousel.datasource = this.NoticiasCarroselInterno;
            carousel.isProduto = this.properties.isProduto;
            this.domElement.innerHTML = "";
            this.domElement.appendChild(carousel);
          }
        }
      );
    } else if (this.properties.Layout == "banner") {
      this.getAllListItemsCarroseInterno().then(
        (ret: NoticiasCarroselInterno): void => {
          this.getBanners().then((item: any): void => {
            if (ret.Carrosel.length > 0 && item) {
              this.NoticiasCarroselInterno = ret;
              const carousel: any = document.createElement(
                "spfx-carousel-banner"
              );
              carousel.datasource = this.NoticiasCarroselInterno;
              carousel.isProduto = this.properties.isProduto;
              carousel.bannerimg = item.SANIconNormal +"?RenditionID=21";
              carousel.bannerurl = item.SANLink;
              carousel.bannertarget = item.SANNewWindown;
              carousel.isDate = this.properties.isDate;
              this.domElement.innerHTML = "";
              this.domElement.appendChild(carousel);
            }
          });
        }
      );
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
  public _NoticiasContent: NoticiasCarrosel = {
    Carrosel: this._Noticias,
    Box1: this._Noticias[0],
    Box2: this._Noticias[1],
  };

  public _NoticiasContentInterno: NoticiasCarroselInterno = {
    Carrosel: this._Noticias,
  };

  public async getAllListItemsCarrosel(): Promise<NoticiasCarrosel> {
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
      const quantidadeCarrosel = quantidade - 2;

      let itemNoticiasCarrosel: PublishingPage[] = [];
      let ItemNoticiaBox1: PublishingPage;
      let ItemNoticiaBox2: PublishingPage;

      // look through the returned items.
      let CountBox1 = 1;
      //const sorted = sort.map((val) => items.find((item) => item.id === val));

      let rFilter = this.properties.multiSelect.map((val) =>
        r.find((item) => item.Id === val)
      );

      for (var i = 0; i < rFilter.length; i++) {
        let iconUrl = null;

        const matches = /SANDestaquePub:SW\|(.*?)\r\n/gi.exec(
          rFilter[i].FieldValuesAsText.MetaInfo
        );
        if (matches !== null && matches.length > 1) {
          // this wil be the value of the PublishingPageImage field
          iconUrl = new SPFXutils().extractIMGUrl(matches[1], "noticias");
        }
        if (i < quantidadeCarrosel) {
          itemNoticiasCarrosel.push({
            Id: rFilter[i].ID,
            Title: rFilter[i].Title,
            Created: rFilter[i].SANDataPublicacao,
            CreatedBy: null,
            SANSinopse1: rFilter[i].SANSinopse1,
            SANAreas: rFilter[i].SANAreas,
            SANAtivo: rFilter[i].SANAtivo,
            SANCategorias: rFilter[i].SANCategorias,
            SANDestaqueHome: rFilter[i].SANDestaqueHome,
            SANOrdem1: rFilter[i].SANOrdem1,
            SANResponsavel: null,
            SANSubTitulo1: rFilter[i].SANSubTitulo1,
            SANFullHtml: rFilter[i].SANSinopse1,
            SANDataVigencia: rFilter[i].SANDataVigencia,
            SANDestaquePub: iconUrl,
            SANDestaqueCarrosel: iconUrl + "?RenditionID=5",
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=7",
            link: rFilter[i].EncodedAbsUrl,
          });
        }
        if (CountBox1 == quantidadeCarrosel + 1) {
          ItemNoticiaBox1 = {
            Id: rFilter[i].ID,
            Title: rFilter[i].Title,
            Created: rFilter[i].SANDataPublicacao,
            CreatedBy: null,
            SANSinopse1: rFilter[i].SANSinopse1,
            SANAreas: rFilter[i].SANAreas,
            SANAtivo: rFilter[i].SANAtivo,
            SANCategorias: rFilter[i].SANCategorias,
            SANDestaqueHome: rFilter[i].SANDestaqueHome,
            SANOrdem1: rFilter[i].SANOrdem1,
            SANResponsavel: null,
            SANSubTitulo1: rFilter[i].SANSubTitulo1,
            SANFullHtml: rFilter[i].SANSinopse1,
            SANDataVigencia: rFilter[i].SANDataVigencia,
            SANDestaquePub: iconUrl,
            SANDestaqueCarrosel: iconUrl + "?RenditionID=5",
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=6",
            link: rFilter[i].EncodedAbsUrl,
          };
        }
        if (CountBox1 == quantidadeCarrosel + 2) {
          ItemNoticiaBox2 = {
            Id: rFilter[i].ID,
            Title: rFilter[i].Title,
            Created: rFilter[i].SANDataPublicacao,
            CreatedBy: null,
            SANSinopse1: rFilter[i].SANSinopse1,
            SANAreas: rFilter[i].SANAreas,
            SANAtivo: rFilter[i].SANAtivo,
            SANCategorias: rFilter[i].SANCategorias,
            SANDestaqueHome: rFilter[i].SANDestaqueHome,
            SANOrdem1: rFilter[i].SANOrdem1,
            SANResponsavel: null,
            SANSubTitulo1: rFilter[i].SANSubTitulo1,
            SANFullHtml: rFilter[i].SANSinopse1,
            SANDataVigencia: rFilter[i].SANDataVigencia,
            SANDestaquePub: iconUrl,
            SANDestaqueCarrosel: iconUrl + "?RenditionID=5",
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=6",
            link: rFilter[i].EncodedAbsUrl,
          };
        }
        CountBox1++;
      }
      this._NoticiasContent.Carrosel = itemNoticiasCarrosel;
      this._NoticiasContent.Box1 = ItemNoticiaBox1;
      this._NoticiasContent.Box2 = ItemNoticiaBox2;

      //verifica se tem noticias
      if (this._NoticiasContent.Box1 === undefined) {
        this._NoticiasContent.Box1 = this._Noticias[0];
      }
      if (this._NoticiasContent.Box2 === undefined) {
        this._NoticiasContent.Box2 = this._Noticias[0];
      }

      return this._NoticiasContent;
    } catch (e) {
      console.error(e);
    }
  }

  public async getAllListItemsCarroseInterno(): Promise<
    NoticiasCarroselInterno
  > {
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

      let itemNoticiasCarrosel: PublishingPage[] = [];

      let rFilter = this.properties.multiSelect.map((val) =>
        r.find((item) => item.Id === val)
      );

      // look through the returned items.
      let CountBox1 = 0;
      for (var i = 0; i < rFilter.length; i++) {
        let iconUrl = null;

        const matches = /SANDestaquePub:SW\|(.*?)\r\n/gi.exec(
          rFilter[i].FieldValuesAsText.MetaInfo
        );
        if (matches !== null && matches.length > 1) {
          // this wil be the value of the PublishingPageImage field
          iconUrl = new SPFXutils().extractIMGUrl(matches[1], "noticias");
        }
        itemNoticiasCarrosel.push({
          Id: rFilter[i].ID,
          Title: rFilter[i].Title,
          Created: this.properties.isProduto
            ? rFilter[i].SANDataInicioComercializacao
            : rFilter[i].SANDataPublicacao,
          CreatedBy: null,
          SANSinopse1: rFilter[i].SANSinopse1,
          SANAreas: rFilter[i].SANAreas,
          SANAtivo: rFilter[i].SANAtivo,
          SANCategorias: rFilter[i].SANCategorias,
          SANDestaqueHome: rFilter[i].SANDestaqueHome,
          SANOrdem1: rFilter[i].SANOrdem1,
          SANResponsavel: null,
          SANSubTitulo1: rFilter[i].SANSubTitulo1,
          SANFullHtml: rFilter[i].SANSinopse1,
          SANDataVigencia: rFilter[i].SANDataVigencia,
          SANDestaquePub: iconUrl,
          SANDestaqueCarrosel: iconUrl + "?RenditionID=13",
          SANDestaqueCarrosel2: iconUrl + "?RenditionID=5",
          link: rFilter[i].EncodedAbsUrl,
          SANFamilia: rFilter[i].SANFamilia,
        });
      }
      this._NoticiasContentInterno.Carrosel = itemNoticiasCarrosel;
      return this._NoticiasContentInterno;
    } catch (e) {
      console.error(e);
    }
  }

  public async getBanners(): Promise<any> {
    try {
      const w = Web(this.properties.siteurlbanner);
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      let camlQuery =
        `<Query><Where><And><Eq><FieldRef Name="SANAtivo" /><Value Type="Boolean">1</Value></Eq><Eq><FieldRef Name="ID" /><Value Type="Counter">` +
        this.properties.bannerid +
        `</Value></Eq></And></Where></Query>`;

      const r = await w.lists
        .getByTitle(this.properties.listnamebanner)
        .renderListDataAsStream({
          ViewXml: `<view>${camlQuery}${queryOptions}</view>`,
        });
      var itemBanner: any;
      // look through the returned items.
      for (var i = 0; i < r.Row.length; i++) {
        itemBanner = {
          Id: r.Row[i].ID,
          Title: r.Row[i].Title,
          SANAtivo: true,
          SANIconNormal: new SPFXutils().extractIMGUrl(
            r.Row[i].SANDestaquePub,
            "banner"
          ),
          SANLink:
            r.Row[i].SANLinkMult == ""
              ? r.Row[i].SANLink
              : r.Row[i].SANLinkMult,
          SANNewWindown: r.Row[i].SANNewWindown,
        };
      }
      return itemBanner;
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
                PropertyPaneTextField("siteurlbanner", {
                  label: strings.SiteUrlbannerLabel,
                }),
                PropertyPaneTextField("listnamebanner", {
                  label: strings.ListNamebannerLabel,
                }),
              ],
            },
            {
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyFieldTermPicker("Tipos", {
                  label: "Selecione o Tipo",
                  panelTitle: "Selecione o Tipo",
                  initialValues: this.properties.Tipos,
                  allowMultipleSelections: true,
                  excludeSystemGroup: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged2,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  limitByGroupNameOrID:
                    "Site Collection - santandernet.sharepoint.com-sites-IPPortugal",
                  limitByTermsetNameOrID: "Tipos",
                  key: "termSetsPickerFieldId",
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
                PropertyFieldOrder("orderedItems", {
                  key: "orderedItems",
                  label: "Ordem",
                  items: this.properties.multiSelect,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                }),
                PropertyPaneTextField("ReadMore", {
                  label: strings.ReadMoreLabel,
                }),
                PropertyPaneToggle("ReadMoreOn", {
                  label: strings.ReadMoreOnLabel,
                }),
                PropertyPaneChoiceGroup("Layout", {
                  label: strings.LayoutLabel,
                  options: [
                    {
                      key: "interno",
                      text: "Carousel Interno",
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
                      key: "2colunas",
                      text: "Carousel Home - 2 colunas",
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
                      key: "banner",
                      text: "Carousel Home - Com Banners",
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
                PropertyPaneSlider("QtdItens", {
                  label: strings.QtdItensLabel,
                  min: 3,
                  max: 12,
                  showValue: true,
                }),
                PropertyPaneTextField("Caml", {
                  label: strings.CamlLabel,
                }),
                PropertyPaneToggle("isProduto", {
                  label: "Produtos e campanhas",
                }),
                PropertyPaneTextField("bannerid", {
                  label: strings.banneridLabel,
                }),
                PropertyPaneToggle("isDate", {
                  label: "Mostrar Data",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
