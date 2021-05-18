import { Fields, ICamlQuery, sp, SPRest } from "@pnp/sp/presets/all";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import { INormalizedResult } from "../dataContracts/INormalizedResult";
import { IPersonValue } from "../dataContracts/IPersonValue";
import * as CamlBuilder from "camljs";
import { calculatePrecision } from "office-ui-fabric-react";
import { IPagePropertyData } from "./IPagePropertyData";
import { odataUrlFrom } from "@pnp/sp/odata";
import { ColumnHelper } from "../util/ColumnHelper";
import { ISelectableLookup } from "../dataContracts/ISelectableLookup";
import { resultItem } from "office-ui-fabric-react/lib/components/FloatingPicker/PeoplePicker/PeoplePicker.scss";

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
    (results.taxonomyPropertyNames.length > 0) ? ColumnHelper.extractTaxonomyValues(results) : "";

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
          jsonValue: ColumnHelper.jsonParseField(result[viewField] || result[viewField + 'Id']),
          personValue: ColumnHelper.extractPersonInfo(htmlValue)
        }
      }

      return normalizedResult;
    });

    for (let key in results.lookupResult) {
      let normalizedLookupResult =  {
        textValue: (typeof results.lookupResult[key]=='object')?results.lookupResult[key].join(","):results.lookupResult[key],
        htmlValue: (typeof results.lookupResult[key]=='object')?results.lookupResult[key].join(","):results.lookupResult[key],
        rawValue: (typeof results.lookupResult[key]=='object')?results.lookupResult[key].join(","):results.lookupResult[key],
        jsonValue: ColumnHelper.jsonParseField(results.lookupResult[key]),
        personValue: null
      };

      (normalizedResults.length==1)?normalizedResults[0][key]=normalizedLookupResult:normalizedResults.push(normalizedLookupResult);
    }

     return normalizedResults;
  }

  /**
  * Returns the Page Properties from current Page
  * @param context : The Webpartcontext to query for correct list and page
  * @param skipSystemFields : Skips all unnecessary System Fields like ID, etc
  */
  public async getPageProperties(context: WebPartContext, skipSystemFields?: boolean): Promise<any> {
    //Make failsafe for Workbench
    let id = (context.pageContext.listItem && context.pageContext.listItem.id) ? context.pageContext.listItem.id : 1;

    const item: any = await sp.web.getList(`${context.pageContext.web.serverRelativeUrl}/SitePages`).items.getById(id).expand("FieldValuesAsText", "FieldValuesAsHtml").get();

    let result = {};
    for (var key in item) {
      let keyLower = key.toLowerCase();
      if (keyLower.substr(0, 1) != "_" && keyLower.search("odata") < 0) {
        result[key] = item[key];
      }
    }
    return PagePropertyService._getValidProperties(result, skipSystemFields);
  }

  /**
  * Returns the Page Properties from current Page
  * @param context : The Webpartcontext to query for correct list and page
  * @param selectableLookup : Fields to query from SharePoint List. Contains expand information. Lookup Columns must be
  *                           expanded to retrieve values
  */
  public async getLookupPropertyValues(context: WebPartContext, selectableLookup: ISelectableLookup): Promise<any> {
    let id = (context.pageContext.listItem && context.pageContext.listItem.id) ? context.pageContext.listItem.id : 1;
    let result = {};
    if (selectableLookup.select.length == 0) { return Promise.resolve(null) };
    return sp.web.getList(`${context.pageContext.web.serverRelativeUrl}/SitePages`).items.getById(id).select(...selectableLookup.select).expand(...selectableLookup.expand).get().then((items) => {

      selectableLookup.expand.forEach((element, index) => {
        if (items[element].length && items[element].length > 0) {
          let tmp = [];
          items[element].forEach(multivalueItem => {
            tmp.push(multivalueItem[selectableLookup.foreignColumns[index]])
          });
          result[element] = tmp;
        } else {
          result[element] = items[element][selectableLookup.foreignColumns[index]];
        }

      });
      return Promise.resolve(result);
    }).catch(error => {
      console.error(error);
      return Promise.resolve(null);
    });
  }

  public async getExpandedPagePropertyValues(context: WebPartContext, propertyNames: string[], selectableLookup: ISelectableLookup): Promise<IPagePropertyData> {
    let expandHelpers: string[] = ["FieldValuesAsText", "FieldValuesAsHtml"];
    //Make failsafe for Workbench
    let id = (context.pageContext.listItem && context.pageContext.listItem.id) ? context.pageContext.listItem.id : 1;
    //return sp.web.lists.getByTitle("Site Pages").items.getById(id).select(propertyNames.join(",")).expand("FieldValuesAsText", "FieldValuesAsHtml").get().then((items) => {
    //    return sp.web.getList(`${context.pageContext.web.serverRelativeUrl}/SitePages`).items.getById(id).select(propertyNames.join(",")).expand("FieldValuesAsText", "FieldValuesAsHtml").get().then((items) => {
    return sp.web.getList(`${context.pageContext.web.serverRelativeUrl}/SitePages`).items.getById(id).select(...propertyNames).expand(...expandHelpers).get().then((items) => {
      var result: IPagePropertyData = { dataItems: (items.length) ? items : [items], taxonomyPropertyNames: [], taxCatchAllResult: null, lookupResult: [] };
      for (var key in items) {
        if (items[key] != null && items[key].WssId != null) {
          result.taxonomyPropertyNames.push(key);
        }
      }

      return this.getLookupPropertyValues(context, selectableLookup).then((lookupItems) => {
        result.lookupResult = lookupItems;
        let itemlist = result.taxonomyPropertyNames.concat(["TaxCatchAll/ID", "TaxCatchAll/Term"]);
        return sp.web.getList(`${context.pageContext.web.serverRelativeUrl}/SitePages`).items.getById(id).select(...itemlist).expand("TaxCatchAll").get().then((termdata) => {
          result.taxCatchAllResult = termdata;
          return Promise.resolve(result);

        }).catch(error => {
          console.info("No Termdata TaxCatchall to fetch:", error);
          result.taxCatchAllResult = null;
          return Promise.resolve(result);
        })
      })
      //return sp.web.lists.getByTitle("Site Pages").items.getById(id).select(itemlist.join(",")).expand("TaxCatchAll").get().then((termdata) => {
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
    return sp.web.getList(`${context.pageContext.web.serverRelativeUrl}/SitePages`).getItemsByCAMLQuery(query, "FieldValuesAsText", "FieldValuesAsHtml").then((result) => {
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


