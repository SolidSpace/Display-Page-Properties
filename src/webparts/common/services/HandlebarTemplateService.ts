import { Text, Log } from '@microsoft/sp-core-library';
import * as strings from 'DisplayPagePropertyWebPartStrings';

export class HandlebarTemplateService {

  public static TEMPLATE_ROWS:string = "Rows";
  public static TEMPLATE_COLUMN:string = "Column";
  public static TEMPLATE_DEBUG:string = "Debug";
  public static TEMPLATE_CUSTOM:string ="Custom";


  public static generateDefaultTemplate(pageProperties: string[]): string {
    let pagePropertyHTML = pageProperties.map((property) => {return Text.format(" <span class= 'ms-DetailsRow-cell'> <strong>{0}: </strong>\{\{{0}.textValue\}\}</span> ", property); }).join("\n");
    let template = Text.format(`<style type="text/css">
      .dynamic-template h2 {
        font-size: 20px;
        font-weight: 600;
        color: "[theme:neutralPrimary, default:#323130]";
      }

      .dynamic-template .dynamic-items .dynamic-item {
          background: "[theme:bodyBackground, default: #fff]";
          border: 1px solid "[theme:neutralLight, default: #edebe9]";
          //box-shadow: 0px 0px 6px #bfbebe;
          margin-bottom: 15px;
      }

      .dynamic-template .dynamic-items .dynamic-item h3 {
          background: "[theme:accentButtonBackground, default:#0078d4]";
          color: "[theme:accentButtonText, default: #fff]";
          padding: 5px 5px 7px 10px;
          margin: 0px;
      }

      .dynamic-template .dynamic-items .dynamic-item .dynamic-item-fields {
          padding: 10px;
      }

      .dynamic-template .dynamic-items .dynamic-item .dynamic-item-fields span {
          display: block;
          font-size: 12px;
      }
  </style>

  <div class="dynamic-template">
      <h2>{0}</h2>
      <div class="dynamic-items">
          {{#each items}}
              <div class="dynamic-item">
                  <h3>Result #{{@index}}</h3>
                  <div class="dynamic-item-fields">
  {1}
                  </div>
              </div>
          {{/each}}
      </div>
  </div>`, strings.DefaultTemplateHeadline,pagePropertyHTML);

    return template;
  }

  public static generateColumnTemplate(pageProperties: string[]):string{
    let pagePropertyHTML = pageProperties.map((property) => {return Text.format("<div class='dynamic-item-field'><div><span class= 'ms-DetailsRow-cell dynamic-item-label'>{0}</span></div><div><span class= 'ms-DetailsRow-cell dynamic-item-value'> <strong>\{\{{0}.textValue\}\} </strong></span></div></div>" , property); }).join("\n");
    let template = Text.format(`<style type="text/css">
    <style type="text/css">
  .dynamic-template h2 {
    font-size: 20px;
    font-weight: 600;
    color: "[theme:neutralPrimary, default:#323130]";
  }

  .dynamic-template .dynamic-items .dynamic-item {
    background: "[theme:bodyBackground, default: #fff]";
    border: 1px solid "[theme:neutralLight, default: #edebe9]";
    //box-shadow: 0px 0px 6px #bfbebe;
    margin-bottom: 15px;
  }

  .dynamic-template .dynamic-items .dynamic-item h3 {
    background: "[theme:accentButtonBackground, default:#0078d4]";
    color: "[theme:accentButtonText, default: #fff]";
    padding: 5px 5px 7px 10px;
    margin: 0px;
  }

  .dynamic-template .dynamic-items .dynamic-item .dynamic-item-fields {
    padding: 10px;
  }

  .dynamic-template .dynamic-items .dynamic-item .dynamic-item-fields span {
    display: block;
    font-size: 12px;
  }
  .dynamic-item{
    display: flex;
  }
  .dynamic-item-field{
    margin:0px 10px 0px 0px;
    display: flex;
  }
  .dynamic-item-fields{
    display: flex;
  }
  .dynamic-item-label {
    font-weight: 800;
    margin: 0px 5px 0px 0px;
    font-size:1.2em;
  }

  .dynamic-item-value {
    font-weight: 300;
  }
  </style>

  <div class="dynamic-template">
    <h2>{0}</h2>
    <div class="dynamic-items">
        {{#each items}}
            <div class="dynamic-item">
                <div class="dynamic-item-fields">
                      {1}
                </div>
            </div>
        {{/each}}
    </div>
  </div>`, strings.DefaultTemplateHeadline,pagePropertyHTML);

    return template;
  }
  public static generateRowTemplate(pageProperties: string[]):string{

    let pagePropertyHTML = pageProperties.map((property) => {return Text.format(" <div  class='dynamic-item-field'><div><span class= 'ms-DetailsRow-cell dynamic-item-label'>{0}</span></div><div><span class= 'ms-DetailsRow-cell dynamic-item-value'> <strong>\{\{{0}.textValue\}\} </strong></span></div></div>", property); }).join("\n");
    let template = Text.format(`<style type="text/css">
    .dynamic-template h2 {
      font-size: 20px;
      font-weight: 600;
      color: "[theme:neutralPrimary, default:#323130]";
    }

    .dynamic-template .dynamic-items .dynamic-item {
        background: "[theme:bodyBackground, default: #fff]";
        border: 1px solid "[theme:neutralLight, default: #edebe9]";
        //box-shadow: 0px 0px 6px #bfbebe;
        margin-bottom: 15px;
    }

    .dynamic-template .dynamic-items .dynamic-item h3 {
        background: "[theme:accentButtonBackground, default:#0078d4]";
        color: "[theme:accentButtonText, default: #fff]";
        padding: 5px 5px 7px 10px;
        margin: 0px;
    }

    .dynamic-template .dynamic-items .dynamic-item .dynamic-item-fields {
        padding: 10px;
    }

    .dynamic-template .dynamic-items .dynamic-item .dynamic-item-fields span {
        display: block;
        font-size: 12px;
    }
    .dynamic-item-field{
      display: flex;
      flex-direction: column;
      margin: 0px 0px 5px 0px;
    }
    .dynamic-item-label{
      font-weight:800;
      margin:0px 5px 0px 0px;
    }
    .dynamic-item-value{
      font-weight:300
    }
  </style>

  <div class="dynamic-template">
    <h2>{0}</h2>
    <div class="dynamic-items">
        {{#each items}}
            <div class="dynamic-item">
                <div class="dynamic-item-fields">
                      {1}
                </div>
            </div>
        {{/each}}
    </div>
  </div>`, strings.DefaultTemplateHeadline,pagePropertyHTML);

    return template;
  }
}
