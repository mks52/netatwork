import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { IODataList } from '@microsoft/sp-odata-types';

import styles from './BlogsWebPart.module.scss';
import * as strings from 'BlogsWebPartStrings';

import MockHttpClient from './MockHttpClient';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment, EnvironmentType} from '@microsoft/sp-core-library';
import {SPComponentLoader} from '@microsoft/sp-loader';
import { IPropertyPaneData } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPane/IPropertyPane';

export interface IBlogsWebPartProps {
  description: string;
  channel: string;
  videoChannel: string;
  theme: string;
  fontface : string;
  headerfontcolor: string;
}

export interface BlogsLists {
  value: BlogsList[];
}

export interface BlogsList {
  Title: string,
  BlogUrl: {Description: string},
  PersonId: number,
  ImageUrl: {Description: string},
  Excerpts: string,
  Display: boolean
}

export default class BlogsWebPartWebPart extends BaseClientSideWebPart<IBlogsWebPartProps> {

  private dropdownOptions: IPropertyPaneDropdownOption[];

  private listsFetched: boolean;
  private fetchLists(url: string) : Promise<any> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
        return null;
      }
    });
  }

  private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
    var url = this.context.pageContext.web.absoluteUrl + `/_api/web/Lists/?$filter=Hidden eq false`;

    return this.fetchLists(url).then((response) => {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        response.value.map((list: IODataList) => {
            options.push( { key: list.EntityTypeName, text: list.Title });
        });
        return options;
    });
  }

  public constructor() {
    super();
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

    SPComponentLoader.loadScript('https://code.jquery.com/jquery-2.2.4.min.js', { globalExportsName: 'jQuery' }).then((jQuery: any): void => {
      SPComponentLoader.loadScript('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js',  { globalExportsName: 'jQuery' }).then((): void => {

      });
    });
  }

  private _renderListAsync(): void {
    //Local environment
    if(Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if(Environment.type === EnvironmentType.SharePoint || EnvironmentType.ClassicSharePoint) {
      this._getListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }

  private _getListData(): Promise<BlogsLists> {
    var url = `https://netatwork212.sharepoint.com/sites/sandbox/_api/web/Lists/HumansofNetAtWorkList/Items`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log(response.json.toString);
        return response.json();

      });
  }

  private _getMockListData(): Promise<BlogsLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
      .then((data:BlogsList[]) =>{
        var listData: BlogsLists = {value: data};
        return listData;
      }) as Promise<BlogsLists>;
  }

  private _renderList(items: BlogsList[]): void {
    let html: string ='';
    items.forEach((item: BlogsList) => {
      if(item.Display){
        html += `
          <div class="col-sm-6 col-md-4">
            <div class="thumbnail">
              <img src="${item.ImageUrl.Description}" alt="${item.Title}">
              <div class="caption">
                <h3>${item.Title}</h3>
                <p>${item.Excerpts}</p>
                <p><a href="${item.BlogUrl.Description}" class="btn btn-default" role="button">Read more</a> </p>
              </div>
            </div>
          </div>
        `;
      }

    });

    const listContainer: Element = this.domElement.querySelector('#spVideoListContainer');
    listContainer.innerHTML = html;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div style="border: 1px solid ${this.properties.theme}" class="ms-bgColor-neutralLighterAlt">
      <div>
        <div style="padding:10px; background-color: ${this.properties.theme}; color: ${this.properties.headerfontcolor}" class="ms-font-xl">
          ${escape(this.properties.description)}
        </div>
      </div>

      <div class="row" id="spVideoListContainer" style="padding: 10px 10px 0px 10px">

      </div>
    </div>
    `;
    this._renderListAsync();


  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName : "List",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Blog Web Part Title"
                })
              ]
            },
          {
            groupName : "Theme",
              groupFields: [
                PropertyPaneDropdown('theme', {
                  label: "Select Color",
                  options: [
                    {key: '#0078d7', text: 'Primary'},
                    {key: '#333333', text: 'Neutral'},
                    {key: '#a6a6a6', text: 'Neutral Tertiary'},
                    {key: '#eaeaea', text: 'Neutral Light'},
                    {key: '#ffb900', text: 'Yellow'},
                    {key: '#e81123', text: 'Red'},
                    {key: '#a80000', text: 'Red Dark'},
                    {key: '#5c2d91', text: 'Purple'},
                    {key: '#00188f', text: 'Blue Mid'},
                    {key: '#008272', text: 'Teal'},
                    {key: '#00B294', text: 'Teal Light'},
                    {key: '#107c10', text: 'Green'},
                    {key: '#bad80a', text: 'Green Light'}
                  ],
                  selectedKey: '#0078d7'
                }),
                PropertyPaneDropdown('headerfontcolor', {
                  label: "Select Font Color for Title",
                  options: [
                    {key: '#ffffff', text: 'White'},
                    {key: '#0078d7', text: 'Primary'},
                    {key: '#333333', text: 'Neutral'},
                    {key: '#a6a6a6', text: 'Neutral Tertiary'},
                    {key: '#eaeaea', text: 'Neutral Light'},
                    {key: '#ffb900', text: 'Yellow'},
                    {key: '#e81123', text: 'Red'},
                    {key: '#a80000', text: 'Red Dark'},
                    {key: '#5c2d91', text: 'Purple'},
                    {key: '#00188f', text: 'Blue Mid'},
                    {key: '#008272', text: 'Teal'},
                    {key: '#00B294', text: 'Teal Light'},
                    {key: '#107c10', text: 'Green'},
                    {key: '#bad80a', text: 'Green Light'}
                  ],
                  selectedKey: '#ffffff'
                }),
                PropertyPaneDropdown('fontface', {
                  label: "Select Font Color for Main Text",
                  options: [
                    {key: '#0078d7', text: 'Primary'},
                    {key: '#333333', text: 'Neutral'},
                    {key: '#a6a6a6', text: 'Neutral Tertiary'},
                    {key: '#eaeaea', text: 'Neutral Light'},
                    {key: '#ffb900', text: 'Yellow'},
                    {key: '#e81123', text: 'Red'},
                    {key: '#a80000', text: 'Red Dark'},
                    {key: '#5c2d91', text: 'Purple'},
                    {key: '#00188f', text: 'Blue Mid'},
                    {key: '#008272', text: 'Teal'},
                    {key: '#00B294', text: 'Teal Light'},
                    {key: '#107c10', text: 'Green'},
                    {key: '#bad80a', text: 'Green Light'}
                  ],
                  selectedKey: '#0078d7'
                })
              ]
          },

          ]
        }
      ]
    };

  }
}
