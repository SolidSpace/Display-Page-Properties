import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneChoiceGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'DisplayPagePropertyWebPartStrings';
import { setup as pnpSetup } from "@pnp/common";
import { sp } from "@pnp/sp/presets/all";
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';
//import { PropertyPaneTextDialogWindow } from '../controls/PropertyPaneTextDialog/PropertyPaneTextDialogWindow';
import { PagePropertyService } from '../common/services/PagePropertyService';
import { PagePropertyConstants } from "../common/constants/PagePropertyConstants";
import { HandlebarTemplateService } from "../common/services/HandlebarTemplateService";
import { get, update } from '@microsoft/sp-lodash-subset';
import { IDisplayPagePropertyWebPartProps } from './IDisplayPagePropertyWebPartProps';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { DisplayPageProperties } from './components/DisplayPageProperties';
import { IDisplayPagePropertiesProps } from './components/IDisplayPagePropertiesProps';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';



export default class DisplayPagePropertyWebPart extends BaseClientSideWebPart<IDisplayPagePropertyWebPartProps> {
  private pagePropertyData: any = [];
  private selectabelPageProperties: IPropertyPaneDropdownOption[];
  private PropertyService: PagePropertyService;
  //private ppTemplateTextDialog: PropertyPaneTextDialogWindow;
  private _themeVariant: IReadonlyTheme | undefined;
  private _themeProvider: ThemeProvider;
  private _defaultPageProperties: string[];
  private _lastProcessedTemplate: string;

  protected onInit(): Promise<void> {

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
    this._defaultPageProperties = ["Title", "ID"];
    this.properties.selectedPageProperties = (this.properties.selectedPageProperties)?this.properties.selectedPageProperties:this._defaultPageProperties;
    this.properties.hasDefaultTemplateBeenUpdated = (this.properties.hasDefaultTemplateBeenUpdated)?this.properties.hasDefaultTemplateBeenUpdated:false;
    this.properties.selectedTemplateLayout = (this.properties.selectedTemplateLayout)?this.properties.selectedTemplateLayout:HandlebarTemplateService.TEMPLATE_DEBUG;

    if(!this.properties.hasDefaultTemplateBeenUpdated){
      switch (this.properties.selectedTemplateLayout) {
        case HandlebarTemplateService.TEMPLATE_COLUMN:
          this.properties.templateContent = HandlebarTemplateService.generateColumnTemplate(this.properties.selectedPageProperties);
          this.properties.hasDefaultTemplateBeenUpdated = false;

          break;
        case HandlebarTemplateService.TEMPLATE_ROWS:
          this.properties.templateContent = HandlebarTemplateService.generateRowTemplate(this.properties.selectedPageProperties);
          this.properties.hasDefaultTemplateBeenUpdated = false;
          break;
        case HandlebarTemplateService.TEMPLATE_DEBUG:
          this.properties.templateContent = HandlebarTemplateService.generateDefaultTemplate(this.properties.selectedPageProperties);
          this.properties.hasDefaultTemplateBeenUpdated = false;
          break;
      }
      this._lastProcessedTemplate = this.properties.selectedTemplateLayout;
    }

    return super.onInit().then(_ => {
      pnpSetup({
        spfxContext: this.context
      });
    }).then(_ => {
      this.PropertyService = new PagePropertyService(sp);
      this.PropertyService.getPageProperties(this.context, false).then((result) => {
        this.pagePropertyData = result;
        this.selectabelPageProperties = this.PropertyService.getSelectableFieldNames(this.pagePropertyData);
      });
    });
  }

  /***************************************************************************
   * Update the current theme variant reference and re-render.
   *
   * @param args The new theme
   ***************************************************************************/
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }


  public render(): void {

    const element: React.ReactElement<IDisplayPagePropertiesProps> = React.createElement(DisplayPageProperties,
      {
        handlebarTemplate: this.properties.templateContent,
        strings: strings.DisplayPagePropertyStrings,
        sp: sp,
        context: this.context,
        skipSystemFields: false,
        themeVariant: this._themeVariant,
        selectedProperties: (this.properties.selectedPageProperties) ? this.properties.selectedPageProperties : this._defaultPageProperties
      }

    );
    ReactDom.render(element, this.domElement);
  }


  public onDisplayModeChanged() {
    this.pagePropertyData = this.PropertyService.onRefreshProperties(this.context, this.displayMode, true, false);
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    if (propertyPath == "selectedTemplateLayout" && oldValue != newValue) {
      if (this._lastProcessedTemplate != this.properties.selectedTemplateLayout) {
        this.properties.selectedPageProperties = (this.properties.selectedPageProperties) ? this.properties.selectedPageProperties : this._defaultPageProperties;

        switch (newValue) {
          case HandlebarTemplateService.TEMPLATE_COLUMN:
            this.properties.templateContent = HandlebarTemplateService.generateColumnTemplate(this.properties.selectedPageProperties);
            this.properties.hasDefaultTemplateBeenUpdated = false;
            break;
          case HandlebarTemplateService.TEMPLATE_ROWS:
            this.properties.templateContent = HandlebarTemplateService.generateRowTemplate(this.properties.selectedPageProperties);
            break;
          case HandlebarTemplateService.TEMPLATE_DEBUG:
            this.properties.templateContent = HandlebarTemplateService.generateDefaultTemplate(this.properties.selectedPageProperties);
            break;
        }
        this._lastProcessedTemplate = this.properties.selectedTemplateLayout;
      }
    }

  }

  public onCustomPropertyPaneChange(propertyPath: string, newValue: any): void {

    update(this.properties, propertyPath, (): any => { return newValue; });
    const oldValue = get(this.properties, propertyPath);
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath == PagePropertyConstants.propertyTemplateText && !this.properties.hasDefaultTemplateBeenUpdated) {
      this.properties.hasDefaultTemplateBeenUpdated = true;
    }
    this.properties.selectedTemplateLayout = HandlebarTemplateService.TEMPLATE_CUSTOM
    this.context.propertyPane.refresh();
  }

  public createDefaultTemplateWithFields(): void {
    update(this.properties,PagePropertyConstants.propertyTemplateContent,():string=>{return HandlebarTemplateService.generateDefaultTemplate(
      get(this.properties,PagePropertyConstants.propertySelectedPageProperties))
    })
//    this.properties.templateContent = HandlebarTemplateService.generateDefaultTemplate(this.properties.selectedPageProperties);
    this.context.propertyPane.refresh();
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Creates a custom PropertyPaneTextDialog for the templateText property
/*
    this.ppTemplateTextDialog = new PropertyPaneTextDialogWindow(PagePropertyConstants.propertyTemplateText, {
      dialogTextFieldValue: this.properties.templateContent,
      onPropertyChange: this.onCustomPropertyPaneChange.bind(this),
      disabled: false,
      strings: strings.TemplateTextStrings
    });
*/
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldMultiSelect('selectedPageProperties', {
                  key: 'multiSelect',
                  label: "Select Properties",
                  options: this.selectabelPageProperties,
                  selectedKeys: this.properties.selectedPageProperties,
                }),
                PropertyPaneChoiceGroup('selectedTemplateLayout', {
                  label: strings.LayoutTemplateLabel,

                  options: [
                    {
                      key: HandlebarTemplateService.TEMPLATE_ROWS,
                      text: strings.LayoutTemplateLabelRows,
                      iconProps: {
                        officeFabricIconFontName: 'GlobalNavButton'
                      }
                    },
                    {
                      key: HandlebarTemplateService.TEMPLATE_COLUMN,
                      text: strings.LayoutTemplateLabelTable,
                      iconProps: {
                        officeFabricIconFontName: 'Table'
                      }
                    },
                    {
                      key: HandlebarTemplateService.TEMPLATE_DEBUG,
                      text: strings.LayoutTemplateLabelDebug,
                      iconProps: {
                        officeFabricIconFontName: 'Code'
                      }
                    },
                    {
                      key: HandlebarTemplateService.TEMPLATE_CUSTOM,
                      text: strings.LayoutTemplateLabelCustom,
                      iconProps: {
                        officeFabricIconFontName: 'CodeEdit'
                      }
                    }

                  ]
                }),
                PropertyFieldCodeEditor('templateContent', {
                  label: strings.CodeEditorButtonLabel,
                  panelTitle: strings.CodeEditorPanelTitle,
                  initialValue: this.properties.templateContent,
                  onPropertyChange: this.onCustomPropertyPaneChange.bind(this),
                  properties: this.properties,
                  disabled: false,
                  key: 'codeEditorFieldId',
                  language: PropertyFieldCodeEditorLanguages.HTML,
                  options: {
                    wrap: true,
                    fontSize: 12
                  }

                })
                //this.ppTemplateTextDialog
              ]
            },

          ]
        }
      ]
    };
  }

}
