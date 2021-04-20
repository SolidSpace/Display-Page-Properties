import * as Handlebars from 'handlebars';
import * as React from 'react';
import { IDisplayPagePropertyTemplateContext } from './IDisplayPagePropertyTemplateContext';
import { Text, Log } from '@microsoft/sp-core-library';
import { PagePropertyService } from '../../common/services/PagePropertyService';
import { IDisplayPagePropertiesProps } from './IDisplayPagePropertiesProps';
import { IDisplayPagePropertiesState } from './IDisplayPagePropertiesState';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'DisplayPagePropertyWebPartStrings';
import { Placeholder }from "@pnp/spfx-controls-react/lib/Placeholder";
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import * as moment from 'moment';


export class DisplayPageProperties extends React.Component<IDisplayPagePropertiesProps, IDisplayPagePropertiesState> {
  private PropertyService: PagePropertyService;
  private templateContext:IDisplayPagePropertyTemplateContext
  private configSet:boolean;

  constructor(props: IDisplayPagePropertiesProps, state: IDisplayPagePropertiesState) {
    super(props);

    let helpers = require<any>('handlebars-helpers')({
      handlebars: Handlebars
    });
    Handlebars.registerHelper("formatdate", (date,format)=> {
      return moment(date).format(format);
    });

    this.state={processedTemplateResult:null,error:null,isError:false,loading:true};
    (this.props.handlebarTemplate=='undefined' || this.props.handlebarTemplate==undefined || this.props.handlebarTemplate=='')?this.configSet=false:this.configSet=true;
    this.PropertyService = new PagePropertyService(this.props.sp);
    this._queryTemplateContent();
  }

  componentDidUpdate(prevProps, prevState){

    if ((prevProps.handlebarTemplate !== this.props.handlebarTemplate)|| (prevProps.selectedProperties.length != this.props.selectedProperties.length) ) {
      this.configSet=true;
      this._queryTemplateContent();
    }
  }


  private async _queryTemplateContent():Promise<any>{
    //let items = await this.PropertyService.getPageProperties(this.props.context,this.props.skipSystemFields);
    this.PropertyService.getExpandedPagePropertyValues(this.props.context,this.props.selectedProperties).then((items)=>{
      let normalizedItems = PagePropertyService.getNormalizedQueryResults(items,this.props.selectedProperties);
      this.templateContext = {items:normalizedItems};
      (this.configSet)?this._renderTemplate():true;
    }).catch(error=>{
      this.setState({error:"<div>"+error+"</div>",isError:true,loading:false});
    });

  }

  private _onConfigure = () => {
    this.props.context.propertyPane.open();
  }

  private _renderTemplate(){
    try{
      let template = Handlebars.compile(this.props.handlebarTemplate);
      let result = template(this.templateContext);

      this.setState({processedTemplateResult:result,isError:false,loading:false,error:null});
    }catch (error) {
      this.setState({processedTemplateResult: null, loading:false, error: Text.format(this.props.strings.errorProcessTemplate, error),isError:true });
    }

  }
  /*************************************************************************************
   * Converts the specified HTML by an object required for dangerouslySetInnerHTML
   * @param html
   *************************************************************************************/
  private createMarkup(html: string) {
    return { __html: this._processTheme(html) };
  }


  private _processTheme = (html: string): string => {
    const { semanticColors, palette } = this.props.themeVariant;

    // Find themable colors
    const expression = /"\[theme:\s?(\w*),\s?default:\s?(.*)]"/;
    let result;

    // For every "[theme:themeVariable, default:defaultColor]" in the template
    while ((result = expression.exec(html)) !== null) {
      // Find the theme variable they're asking for
      const themeVariable: string = result[1];

      // Find the theme equivalent or default value
      const themeColor = (semanticColors[themeVariable] ? semanticColors[themeVariable] : palette[themeVariable] ? palette[themeVariable] : result[2]);

      // Replace the color
      html = html.replace(result[0], themeColor);
    }

    return html;
  }

  public render(): React.ReactElement<IDisplayPagePropertiesProps> {
    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;
    if(!this.configSet){
      return(
      <Placeholder iconName='Edit'
      iconText={strings.ConfigurePageProperties}
      description={this.props.strings.propmessage}
      buttonLabel={strings.PlaceholderButtonLabel}
      onConfigure={this._onConfigure} >
      </Placeholder>
      );
    }else{
      if(this.state.loading){
      }else{
        <Spinner label={this.props.strings.loading} ariaLive="assertive" labelPosition="right" />
      }
      return (
        <div style={{backgroundColor: semanticColors.bodyBackground, color: semanticColors.bodyText}}>
          <div dangerouslySetInnerHTML={(this.state.isError)?this.createMarkup(this.state.error):this.createMarkup(this.state.processedTemplateResult)}></div>
        </div>

      );
    }
  }
}


