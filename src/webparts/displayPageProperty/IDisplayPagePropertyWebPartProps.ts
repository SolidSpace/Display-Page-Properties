import { ILookupColumn } from "../common/dataContracts/ILookupColumn";

export interface IDisplayPagePropertyWebPartProps {
  skipSystemFields: string;
  templateContent: string;
  hasDefaultTemplateBeenUpdated:boolean;
  selectedPageProperties:string[];
  selectedTemplateLayout:string;
  selectedLookupProperties: ILookupColumn[];
}
