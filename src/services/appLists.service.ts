import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";
import {
  PublishingPage,
  NoticiasCarrosel
} from "./../interfaces/appLists.interface";
import { SPFXutils } from "./util.service";
export class NoticiasDS {
  constructor() {}

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
      SANDestaqueCarrosel2: "iconUrl"
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
      SANDestaqueCarrosel2: "iconUrl"
    }
  ];
  public _NoticiasContent: NoticiasCarrosel = {
    Carrosel: this._Noticias,
    Box1: this._Noticias,
    Box2: this._Noticias
  };

  public async _prodGetNoticasInfo(
    siteurl: string,
    quantidade: number,
    fieldsFilter: string
  ): Promise<NoticiasCarrosel> {
    try {
      const w = Web(siteurl);
      const rowLimit = `<RowLimit>${quantidade}</RowLimit>`;
      const queryOptions = `<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>`;
      const camlQuery = `<Query><Where><And><Eq><FieldRef Name='SANDestaquePrincipal' /><Value Type='Boolean'>1</Value></Eq><Eq><FieldRef Name='SANAtivo' /><Value Type='Boolean'>1</Value></Eq></And></Where><OrderBy><FieldRef Name='SANOrdem1' Ascending='True' /></OrderBy></Query>`;
      const r = await w.lists.getByTitle("Pages").getItemsByCAMLQuery(
        {
          ViewXml: `<View>${camlQuery}${queryOptions}${rowLimit}</View>`
        },
        "FieldValuesAsText"
      );
      const quantidadeCarrosel = quantidade - 2;
      var itemNoticias: NoticiasCarrosel;
      let itemNoticiasCarrosel: PublishingPage[] = [];
      let ItemNoticiaBox1: PublishingPage[] = [];
      let ItemNoticiaBox2: PublishingPage[] = [];
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

        // var item: PublishingPage = {
        //   Id: r[i].ID,
        //   Title: r[i].Title,
        //   Created: r[i].Created,
        //   CreatedBy: null,
        //   SANSinopse1: r[i].SANSinopse1,
        //   SANAreas: r[i].SANAreas,
        //   SANAtivo: r[i].SANAtivo,
        //   SANCategorias: r[i].SANCategoriasd,
        //   SANDestaqueHome: r[i].SANDestaqueHome,
        //   SANOrdem1: r[i].SANOrdem1,
        //   SANResponsavel: null,
        //   SANSubTitulo1: r[i].SANSubTitulo1,
        //   SANFullHtml: r[i].SANSinopse1,
        //   SANDataVigencia: r[i].SANDataVigencia,
        //   SANDestaquePub: iconUrl
        // };
        if (i < quantidadeCarrosel) {
          itemNoticiasCarrosel.push({
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
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=6"
          });
          if (i == quantidadeCarrosel - 1) {
            CountBox1++;
          }
        }
        if (CountBox1 == 1) {
          ItemNoticiaBox1.push({
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
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=6"
          });
        }
        if (CountBox1 == 2) {
          ItemNoticiaBox2.push({
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
            SANDestaqueCarrosel2: iconUrl + "?RenditionID=6"
          });
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

  public async GetNoticiasCarrosel(
    siteUrl: string,
    quantidade: number,
    fieldsFilter?: string
  ): Promise<NoticiasCarrosel> {
    return this._prodGetNoticasInfo(siteUrl, quantidade, fieldsFilter);
  }
}
