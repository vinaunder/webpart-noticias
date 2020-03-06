export class SPFXutils {
  public extractIMGUrl(urlFoto?: string, tipo?: string): string {
    let iconUrl = null;
    if (
      !urlFoto ||
      urlFoto == undefined ||
      urlFoto == "" ||
      urlFoto.length == 0 ||
      urlFoto.match(/<img/) == null
    ) {
      iconUrl = "/sites/IPPortugal/PublishingImages/" + tipo + "semfoto.jpg";
      return iconUrl;
    }
    iconUrl = /src\s*=\s*"(.+?)"/.exec(urlFoto)[1];
    return iconUrl;
  }
  public extractStrings(strings: string[]) {}
}
