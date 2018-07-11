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

import styles from './NewsWebPart.module.scss';
import * as strings from 'NewsWebPartStrings';

import MockHttpClient from './MockHttpClient';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment, EnvironmentType} from '@microsoft/sp-core-library';
import {SPComponentLoader} from '@microsoft/sp-loader';



export interface INewsWebPartProps {
  description: string;
  theme: string;
  fontface: string;
  headerfontcolor: string;
}

export interface NewsLists {
  value: NewsList[];
}

export interface NewsList {
  Title: string,
  Image: {Url: string},
  Link: {Url: string},
  Description: string,
  Display: boolean
}
export default class NewsWebPartWebPart extends BaseClientSideWebPart<INewsWebPartProps> {
    private dropdownOptions: IPropertyPaneDropdownOption[];

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

  private _getListData(): Promise<NewsLists> {
    var url = `https://netatwork212.sharepoint.com/sites/sandbox/_api/web/Lists/NewsList/Items`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log(response.json.toString);
        return response.json();

      });
  }

  private _getMockListData(): Promise<NewsLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
      .then((data:NewsList[]) =>{
        var listData: NewsLists = {value: data};
        return listData;
      }) as Promise<NewsLists>;
  }

  private _renderList(items: NewsList[]): void {
    let i : number =0;
    let html1: string ='';
    let html2: string ='';


    items.forEach((item: NewsList) => {
      if(item.Display){
        if(i == 0) {
          html1 += `<li data-target="#carousel-example-generic" data-slide-to="0" class="active"></li>`;
          html2 += `
          <div class="item active">
            <img src="${item.Image.Url}" alt="${item.Title}">
            <div>
              <a href="${item.Link.Url}" target="_new">
              <div style="padding-top:10px; padding-bottom:10px; padding-left: 15px; padding-rigth:15px" class="bg-primary">
                <h4>${item.Title}</h4>
                ${item.Description}
              </div>
              </a>
            </div>
          </div>`;
        } else {
          html1 += `<li data-target="#carousel-example-generic" data-slide-to="${i}"></li>`;
          html2 += `
          <div class="item ms-sm12">
            <img src="${item.Image.Url}" alt="${item.Title}">
            <div>
              <a href="${item.Link.Url}" target="_new">
              <div style="padding-top:10px; padding-bottom:10px; padding-left: 15px; padding-rigth:15px" class="bg-primary">
                <h4>${item.Title}</h4>
                ${item.Description}
              </div>
              </a>
            </div>
          </div>`;
        }

      i++;
      }
    });

    const listContainer1: Element = this.domElement.querySelector('#ol1');
    const listContainer2: Element = this.domElement.querySelector('#listbox1');
    // listContainer1.innerHTML = html1;
    listContainer2.innerHTML = html2;
  }

  public render(): void {
    this.domElement.innerHTML = `

    <div id="carousel-example-generic" class="carousel slide" data-ride="carousel">
      <!-- Indicators -->
      <div style="border: 1px solid ${this.properties.theme}" class="ms-bgColor-neutralLighterAlt" >
      <div>
        <div style="padding:10px; background-color: ${this.properties.theme}; color: ${this.properties.headerfontcolor}" class="ms-font-xl">
          ${escape(this.properties.description)}
        </div>
      </div>

      <!-- Wrapper for slides -->
      <div class="carousel-inner" role="listbox" id="listbox1">
        ...
      </div>

      <!-- Controls -->
      <a class="left carousel-control" href="#carousel-example-generic" role="button" data-slide="prev">
        <span class="glyphicon glyphicon-chevron-left" aria-hidden="true"></span>
        <span class="sr-only">Previous</span>
      </a>
      <a class="right carousel-control" href="#carousel-example-generic" role="button" data-slide="next">
        <span class="glyphicon glyphicon-chevron-right" aria-hidden="true"></span>
        <span class="sr-only">Next</span>
      </a>
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
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
            ]
          },
          ]
        }
      ]
    };
  }
}
