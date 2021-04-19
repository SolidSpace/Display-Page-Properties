import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPRest } from "@pnp/sp";
import { IDisplayPagePropertyStrings } from "./IDisplayPagePropertiesStrings";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IDisplayPagePropertiesProps {
  handlebarTemplate:string;
  strings:IDisplayPagePropertyStrings;
  sp:SPRest;
  context:WebPartContext;
  skipSystemFields?:boolean;
  themeVariant: IReadonlyTheme;
  selectedProperties:string[];
}
