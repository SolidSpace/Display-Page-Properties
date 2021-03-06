# sp-page-properties

## What is the purpose of the Webpart?
Adding SharePoint columns to the SitePages Library is a common way to set up filter and searchable Content Sites with Modern Search Webparts. 
This "Page Property" Webpart allows you to render those additional columns within your Page and put them in shape by setting up an awesome handlebars template.


This Webpart is build for SharePoint Online with spfx version 1.11.

![Web Part Preview](https://raw.githubusercontent.com/SolidSpace/Display-Page-Properties/main/assets/display-page-properties-webpart.GIF)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)
![Node.js LTS 10.x](https://img.shields.io/badge/Node.js-LTS%2010.x-green.svg)
![SharePoint Online](https://img.shields.io/badge/SharePoint-Online-red.svg)
![Teams N/A](https://img.shields.io/badge/Teams-N%2FA-lightgrey.svg)
![Workbench Hosted](https://img.shields.io/badge/Workbench-Hosted-yellow.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Solution

Solution|Author(s)
--------|---------
displayPageProperty| Marc Jonas

## Version history

Version|Date|Comments
-------|----|--------
1.0|March 19, 2021|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**
- For debug purpose, disable following section in gulp conf, to ensure that breakpoint will work correctly. 
 ```gulp
  generatedConfiguration.module.rules.push(
        { test: /\.js$/, loader: 'unlazy-loader' }
      );
```

## Building the code

Download & install all dependencies, build, bundle & package the project

## download & install dependencies
```shell
npm install
```

## transpile all TypeScript & SCSS => JavaScript & CSS
```shell
gulp build
```

## create component bundle & manifest
```shell
gulp bundle
```

## create SharePoint package
```shell
gulp package-solution
```

These commands produce the following:

- `./lib`: intermediate-stage commonjs build artifacts
- `./dist`: bundled script, along with other resources
- `./temp/deploy`: all resources required by component(s) to deploy to a CDN (when `--ship` argument present)

## Build options

- `gulp clean`: Deletes all build output (`/dist`, `/lib`, `/temp`, etc.).
- `gulp build`: Transpiles all TypeScript & SCSS to JavaScript & CSS, generates source map files & TypeScript type declaration files
- `gulp bundle [--ship|-p|--production]`: Runs gulp task **build**, then uses webpack to create the JavaScript bundle(s) and component manifest(s) as defined in `./config/config.json`. The `--ship`, `-p` or `--production` argument specifies a production build that will generate minified bundles.
- `gulp serve [--ship|-p|--production]`: Runs gulp tasks `build`, `bundle` & starts the local webserver. Depending on the project type, it opens the browser and navigates to the local workbench or specified URL (in the case of extension components). The `--ship`, `-p` or `--production` argument specifies a production build that modifies the resulting package for production hosting rather than local hosting of assets.
- `gulp package-solution`: Creates the SharePoint Package (`.sppkg`) file.
- `gulp dist`: Creates a production-ready SharePoint Package (`.sppkg`) file. The following gulp task gets executed in this specific order `gulp clean`, `gulp bundle`, `gulp package-solution.`
- `gulp dev`: Creates a development-ready SharePoint Package (`.sppkg`) file. The following gulp task will be executed in this specific order `gulp clean`, `gulp bundle`, `gulp package-solution.`

> View all available gulp tasks by running `gulp --tasks`


## Credits 
Adopted code from the SPFX Project "Content Query Online" that helped me to understand, how to get 
Handlebars in SPFX up and running and fix loading issues with handlebar-helpers. 
Thanks a lot to the people who build the content query webpart. 
https://github.com/pnp/sp-dev-fx-webparts/tree/master/samples/react-content-query-online

## Features

#### Displaying items and their values

For displaying items and their field values, we must first iterate through the exposed `{{items}}` token using a `{{each}}` block helper:
To display items and their field values, you have to iterate through the {{items}} token by using a each loop from the handlebars block helpers:

##### Handlebars

```handlebars
{{#each items}}
    <div class="item"></div>
{{/each}}
```
Before customizing your template. Select your desired fields from the "Selected Page properties" box. 

##### Handlebars

```handlebars
{{#each items}}
    <div class="item">
        <p>MyField value : {{MyField}}</p>
    </div>
{{/each}}
```

The above code will render an `[object]`. To access the field value, the following method are available.

Property | Description
---------|---------------
`{{MyField.textValue}}` | Renders the text value of the field, a more readable end-user value to use for display.
`{{MyField.htmlValue}}` | Renders the HTML value of the field. For example, a *Link* field HTML value would render something like `<a href="...">My Link Field</a>`
`{{MyField.rawValue}}`  | Returns the raw value of the field. For example, a *Taxonomy* field raw value would return an object which contains the term `wssId` ,its label and a term property with the correct term. Please note, that a multivalue term field will return an array ob term objects. Use #each block helper to render
`{{MyField.jsonValue}}`  | Returns a JSON object value of the field. For example, an *Image* field JSON value would return a JSON object which contains the `serverRelativeUrl` property
`{{MyField.personValue}}`  | Returns an object value of a person field. The `personValue` property provides `email`, `displayName` and `image` properties. The `image` property contains `small`, `medium`, and `large` properties, each of which pointing to the profile image URL for the small, medium, and large profile images.


##### Handlebars

```handlebars
{{#each items}}
    <div class="item">
        <p>MyField text value : {{MyField.textValue}}</p>
        <p>MyField html value : {{MyField.htmlValue}}</p>
        <p>MyField raw value : {{MyField.rawValue}}</p>
        <p>MyImageField JSON value : {{MyImageField.jsonValue}}</p>
        <p>MyPersonField person value : {{MyPersonField.personValue}}</p>
    </div>
{{/each}}
```

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
