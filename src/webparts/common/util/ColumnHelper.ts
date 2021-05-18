import { IPersonValue } from "../dataContracts/IPersonValue";
import { IPagePropertyData } from "../services/IPagePropertyData";

export class ColumnHelper {
  /**************************************************************************************************
* Returns a JSON object
* @param value : A string representation of a JSON object
**************************************************************************************************/
  public static jsonParseField(value: any): any {
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
   public static extractPersonInfo(htmlValue: string): IPersonValue {

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
   public static extractTaxonomyValues(results: IPagePropertyData): IPagePropertyData {
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
          let term = taxData[listItem[key].WssId];
          let newTerm = { ...listItem[key], "Term": term };
          listItem[key] = newTerm;
        }
      }
    });
    return results;
  }

}


