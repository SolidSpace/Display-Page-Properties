import { Fields, ICamlQuery, sp, SPRest } from "@pnp/sp/presets/all";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import { INormalizedResult } from "../dataContracts/INormalizedResult";
import { IPersonValue } from "../dataContracts/IPersonValue";
import * as CamlBuilder from "camljs";
import { calculatePrecision } from "office-ui-fabric-react";
import { IPagePropertyData } from "./IPagePropertyData";

export class PagePropertyService {
  private sp: SPRest

  private static _systemColumns: Array<string> = ["FileSystemObjectType", "Id", "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl", "ContentTypeId", "ComplianceAssetId", "WikiField", "BannerImageUrl", "PromotedState", "FirstPublishedDate", "ID", "Created", "AuthorId", "Modified", "EditorId", "CheckoutUserId", "GUID", "CanvasContent1", "LayoutWebpartsContent"];
  private static _unsafeColumns: Array<string> = ["CanvasContent1", "LayoutWebpartsContent", "ComplianceAssetId", "FileSystemObjectType", "ServerRedirectEmbedUri", "WikiField"];
  private static _ignoreColumns: Array<string> = ["FieldValuesAsText", "FieldValuesAsHtml"];

  constructor(sp: SPRest) {
    this.sp = sp;
  }

  /**************************************************************************************************
   * Normalizes the results coming from a CAML query into a userfriendly format for handlebars
   * @param results : The results returned by a CAML query executed against a list
   **************************************************************************************************/
  public static getNormalizedQueryResults(results: IPagePropertyData, viewFields: string[]): INormalizedResult[] {
    //Map Taxonomy Terms to Taxonomy Fields
    (results.taxonomyPropertyNames.length > 0) ? PagePropertyService.extractTaxonomyValues(results) : "";

    const normalizedResults: INormalizedResult[] = results.dataItems.map((result) => {
      let normalizedResult: any = {};
      let formattedCharsRegex = /_x00(20|3a|[c-f]{1}[0-9a-f]{1})_/gi;
      for (let viewField of viewFields) {

        //check if the intenal fieldname begins with a special character (_x00)
        let viewFieldOdata: string = viewField;
        if (viewField.indexOf("_x00") == 0) {
          viewFieldOdata = `OData_${viewField}`;
        }
        let formattedName = viewFieldOdata.replace(formattedCharsRegex, "_x005f_x00$1_x005f_");
        formattedName = formattedName.replace(/_x00$/, "_x005f_x00");
        let isTerm: boolean = (results.taxonomyPropertyNames.indexOf(viewField) > -1) ? true : false;

        const htmlValue: string = result.FieldValuesAsHtml[formattedName];
        normalizedResult[viewField] = {
          textValue: (isTerm) ? result[viewField].Term : result.FieldValuesAsText[formattedName],
          htmlValue: (isTerm) ? result[viewField].Term : htmlValue,
          rawValue: result[viewField] || result[viewField + 'Id'],
          jsonValue: this.jsonParseField(result[viewField] || result[viewField + 'Id']),
          personValue: this.extractPersonInfo(htmlValue)
        }
      }

      return normalizedResult;
    });

    return normalizedResults;
  }



  /**************************************************************************************************
  * Returns a JSON object
  * @param value : A string representation of a JSON object
  **************************************************************************************************/
  private static jsonParseField(value: any): any {
    if (typeof value === 'string') {
      try {
        let jsonObject = JSON.parse(value);
        return jsonObject;
      }
      catch {
        return value;
      }
    }
    return value;
  }

  /**
    * Returns user profile information based on a user field
    * @param htmlValue : A string representation of the HTML field rendering
    * This function does a very rudimentary extraction of user information based on very limited
    * HTML parsing. We need to update this in the future to make it more sturdy.
    */
  private static extractPersonInfo(htmlValue: string): IPersonValue {

    try {
      const sipIndex = htmlValue.indexOf(`sip='`);
      if (sipIndex === -1) {
        return null;
      }
      // Try to extract the user email and name

      // Get the email address -- we should use RegExp for this, but I suck at RegExp
      const sipValue = htmlValue.substring(sipIndex + 5, htmlValue.indexOf(`'`, sipIndex + 5));
      const anchorEnd: number = htmlValue.lastIndexOf('</a>');
      const anchorStart: number = htmlValue.substring(0, anchorEnd).lastIndexOf('>');
      const name: string = htmlValue.substring(anchorStart + 1, anchorEnd);

      // Generate picture URLs
      const smallPictureUrl: string = `/_layouts/15/userphoto.aspx?size=S&username=${sipValue}`;
      const medPictureUrl: string = `/_layouts/15/userphoto.aspx?size=M&username=${sipValue}`;
      const largePictureUrl: string = `/_layouts/15/userphoto.aspx?size=L&username=${sipValue}`;


      let result: IPersonValue = {
        email: sipValue,
        displayName: name,
        picture: {
          small: smallPictureUrl,
          medium: medPictureUrl,
          large: largePictureUrl
        }
      };
      return result;

    } catch (error) {

      return null;
    }
  }
  /**************************************************************************************************
  * Default return value for getListItems does not contain the Term for Taxonmy Fields.
  * This function maps the Terms to the Taxonmy Fields.
  * @param results : Queried ListItem Data from getExpandedPagePropertyValues
  **************************************************************************************************/
  private static extractTaxonomyValues(results: IPagePropertyData): IPagePropertyData {
    let taxData = {};

    //Build Term ID Object for easier access
    results.taxCatchAllResult.TaxCatchAll.forEach(taxItem => {
      let id = taxItem.ID;
      taxData = { ...taxData, [id]: taxItem.Term };
    });

    //cycle each returned listitem object
    results.dataItems.forEach((listItem) => {
      for (var key in listItem) {
        //is field recognized as taxonomy field
        if (results.taxonomyPropertyNames.indexOf(key) > -1) {
          let term =taxData[listItem[key].WssId];
          let newTerm = { ...listItem[key], "Term": term };
          listItem[key] = newTerm;
        }
      }
    });
    return results;
  }

  /**
  * Returns the Page Properties from current Page
  * @param context : The Webpartcontext to query for correct list and page
  * @param skipSystemFields : Skips all unnecessary System Fields like ID, etc
  */
  public async getPageProperties(context: WebPartContext, skipSystemFields?: boolean): Promise<any> {
    //Make failsafe for Workbench
    let id = (context.pageContext.listItem && context.pageContext.listItem.id) ? context.pageContext.listItem.id : 1;
    const item: any = await sp.web.lists.getByTitle("Site Pages").items.getById(id).expand("FieldValuesAsText", "FieldValuesAsHtml").get();

    let result = {};
    for (var key in item) {
      let keyLower = key.toLowerCase();
      if (keyLower.substr(0, 1) != "_" && keyLower.search("odata") < 0) {
        result[key] = item[key];
      }
    }
    return PagePropertyService._getValidProperties(result, skipSystemFields);
  }

  public async getExpandedPagePropertyValues(context: WebPartContext, propertyNames: string[]): Promise<IPagePropertyData> {
    let expandHelpers: string[] = ["FieldValuesAsText", "FieldValuesAsHtml"];
    let expandAll: string[] = expandHelpers.concat(propertyNames);
    //Make failsafe for Workbench
    let id = (context.pageContext.listItem && context.pageContext.listItem.id) ? context.pageContext.listItem.id : 1;
    return sp.web.lists.getByTitle("Site Pages").items.getById(id).select(propertyNames.join(",")).expand("FieldValuesAsText", "FieldValuesAsHtml").get().then((items) => {
      var result: IPagePropertyData = { dataItems: (items.length) ? items : [items], taxonomyPropertyNames: [], taxCatchAllResult: null };
      for (var key in items) {
        if (items[key] != null && items[key].WssId != null) {
          result.taxonomyPropertyNames.push(key);
        }
      }
      let itemlist = result.taxonomyPropertyNames.concat(["TaxCatchAll/ID", "TaxCatchAll/Term"]);
      return sp.web.lists.getByTitle("Site Pages").items.getById(id).select(itemlist.join(",")).expand("TaxCatchAll").get().then((termdata) => {
        result.taxCatchAllResult = termdata;
        return Promise.resolve(result);
      }).catch(error => {
        console.info("No Termdata TaxCatchall to fetch:", error);
        result.taxCatchAllResult = null;
        return Promise.resolve(result);
      });
    }).catch(error => {
      console.error("Error fetching Propertydata:", error);
      return Promise.reject(error);
    });


  }

  /**
   * CAML Query integratgion - Currently not used. may be used in later versions
   * @param context : The Webpartcontext to query for correct list and page
   */
  public async getPropertyDataByCAML(context: WebPartContext, caml: string): Promise<any> {
    //Make failsafe for Workbench
    let id = (context.pageContext.listItem && context.pageContext.listItem.id) ? context.pageContext.listItem.id : 1;
    var camlBuilder = new CamlBuilder();
    //var caml:string =  camlBuilder.Where().NumberField(id).EqualTo(context.pageContext.listItem.id).ToString();
    const query: ICamlQuery = {
      ViewXml: caml
    }
    return sp.web.lists.getByTitle("Site Pages").getItemsByCAMLQuery(query, "FieldValuesAsText", "FieldValuesAsHtml").then((result) => {
      return Promise.resolve(result);
    }).catch((error) => {
      console.error(error);
      return Promise.reject(error);
    })
  }

  public onRefreshProperties(context: WebPartContext, displayMode: DisplayMode, forceRefresh?: boolean, skipSystemFields?: boolean): any {
    if (displayMode == DisplayMode.Read || forceRefresh) {
      this.getPageProperties(context, skipSystemFields).then((result) => {
        return result;
      }).catch((error) => {
        console.error(error);
      });
    }
  }

  /**
   * Extracts an Array of Fieldnames from the pageProperties
   * @param pageProperties :  SharePoint result from getPageProperties
   */
  public getFieldNamesFromProperties(pageProperties: any): string[] {
    let result: string[] = [];
    for (var key in pageProperties) {
      if (key != "FieldValueAsText" && key != "FieldValuesAsHtml") {
        result.push(key)
      }
    }
    return result;
  }

  /**
   * Extracts an Key/Value Array of Fieldnames from the pageProperties.
   * Used in PropertyPane Field Selection
   * @param pageProperties :  SharePoint result from getPageProperties
   */
  public getSelectableFieldNames(pageProperties: any): IPropertyPaneDropdownOption[] {
    let result: IPropertyPaneDropdownOption[] = [];
    for (var key in pageProperties) {
      if (PagePropertyService._ignoreColumns.indexOf(key) == -1) {
        result.push({ "key": key, "text": key });
      }
    }
    return result;
  }

  /**
   * Some Default Page Properties cause issues. This Method makes them unavailable to the user
   * @param pageProperties :  SharePoint result from getPageProperties
   */
  private static _getValidProperties(pageProperties: any, skipSystemFields: boolean): any {

    if (pageProperties.Title != undefined) {
      let skipColumns: Array<string> = (skipSystemFields) ? PagePropertyService._unsafeColumns.concat(...PagePropertyService._systemColumns) : PagePropertyService._unsafeColumns;
      skipColumns.forEach((keyName) => {
        delete pageProperties[keyName];
      });
    }
    return pageProperties;

  }


}

