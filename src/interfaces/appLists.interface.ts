export interface PublishingPage {
  Id: number;
  Title: string;
  Created: string;
  CreatedBy: string;
  SANSubTitulo1: string;
  SANSinopse1: string;
  SANResponsavel: string;
  SANOrdem1: number;
  SANDestaquePub: string;
  SANDestaqueCarrosel: string;
  SANDestaqueCarrosel2: string;
  SANDestaqueHome: boolean;
  SANDataVigencia: Date;
  SANAtivo: boolean;
  SANAreas: string[];
  SANFullHtml: string;
  SANCategorias: string[];
  link: string;
}
export interface NoticiasCarrosel {
  Carrosel: PublishingPage[];
  Box1: PublishingPage;
  Box2: PublishingPage;
}
export interface NoticiasCarroselInterno {
  Carrosel: PublishingPage[];
}

export interface NoticiasHome {
  Box1: PublishingPage[];
  Box2: PublishingPage;
}
