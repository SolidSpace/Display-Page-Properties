import { ILookupColumn } from "../dataContracts/ILookupColumn";
import { ISelectableLookup } from "../dataContracts/ISelectableLookup";

export class LookupHelper{

  public static getSelectableLookup(lookupProperties:ILookupColumn[]):ISelectableLookup{
    let result:ISelectableLookup={expand:[],select:[],foreignColumns:[]};
    if(lookupProperties){
      lookupProperties.forEach(property=>{
        result.select.push(`${property.lookupColumn}/${property.foreignColumn}`);
        result.expand.push(property.lookupColumn);
        result.foreignColumns.push(property.foreignColumn);
      });
    }

    return result;
  }

}
