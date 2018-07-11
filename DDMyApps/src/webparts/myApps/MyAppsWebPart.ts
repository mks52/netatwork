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

import styles from './MyAppsWebPart.module.scss';
import * as strings from 'MyAppsWebPartStrings';

import MockHttpClient from './MockHttpClient';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment, EnvironmentType} from '@microsoft/sp-core-library';


export interface IMyAppsWebPartProps {
  description: string;
  theme: string;
  fontface: string;
  headerfontcolor: string;
}

export interface MyAppsLists {
  value: MyAppsList[];
}

export interface MyAppsList {
  Title: string,
  TypeofApp: string,
  O365App: string,
  ThirdPartyAppImageLink: {Url: string},
  ThirdPartyAppUrl: {Url: string}
}

export default class MyAppsWebPartWebPart extends BaseClientSideWebPart<IMyAppsWebPartProps> {
private dropdownOptions: IPropertyPaneDropdownOption[];

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

  private _getListData(): Promise<MyAppsLists> {
    var url = `https://netatwork212.sharepoint.com/sites/sandbox/_api/web/Lists/MyAppsList/Items`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log(response.json.toString);
        return response.json();
      });
  }

  private _getMockListData(): Promise<MyAppsLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
      .then((data:MyAppsList[]) =>{
        var listData: MyAppsLists = {value: data};
        return listData;
      }) as Promise<MyAppsLists>;
  }

  private _renderList(items: MyAppsList[]): void {
    let html: string ='';
    let o365logo: string[] =[];
    o365logo["Delve"] = "DelveLogo";
    o365logo["Dynamics 365"] = "Dynamics365Logo";
    o365logo["Excel"] = "ExcelLogo";
    o365logo["Forms"] = "OfficeFormsLogo";
    o365logo["OneDrive"] = "OneDrive";
    o365logo["OneNote"] = "OneNoteLogo";
    o365logo["Outlook"] = "OutlookLogo";
    o365logo["People"] = "People";
    o365logo["PowerPoint"] = "PowerPointLogo";
    o365logo["SharePoint"] = "SharepointLogo";
    o365logo["Sway"] = "SwayLogo";
    o365logo["Tasks"] = "TaskLogo";
    o365logo["Teams"] = "TeamsLogo";
    o365logo["Videos"] = "OfficeVideoLogo";
    o365logo["Word"] = "WordLogo";
    o365logo["Yammer"] = "YammerLogo";
    o365logo["Delvecolor"] = "blue";
    o365logo["Dynamics 365color"] = "blueDark";
    o365logo["Excelcolor"] = "greenDark";
    o365logo["Formscolor"] = "teal";
    o365logo["OneDrivecolor"] = "blue";
    o365logo["OneNotecolor"] = "purple";
    o365logo["Outlookcolor"] = "blue";
    o365logo["Peoplecolor"] = "blue";
    o365logo["PowerPointcolor"] = "orange";
    o365logo["SharePointcolor"] = "blue";
    o365logo["Swaycolor"] = "teal";
    o365logo["Taskscolor"] = "blue";
    o365logo["Teamscolor"] = "purple";
    o365logo["Videoscolor"] = "neutralTertiary";
    o365logo["Wordcolor"] = "blueMid";
    o365logo["Yammercolor"] = "blue";

    items.forEach((item: MyAppsList) => {
      html += `
        <div style="padding-bottom:10px" class="ms-Grid-col ms-sm4 ms-md3 ms-lg2">
          <a href="${item.ThirdPartyAppUrl.Url}" target="_blank">
          <div class="ms-Grid-row">`;
          if(item.TypeofApp == "O365"){
            html += `
            <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-fontColor-${o365logo[item.O365App+"color"]}" style="text-align:center">
              <i class="ms-Icon ms-Icon--${o365logo[item.O365App]}" style="font-size:xx-large"  aria-hidden="true"></i>
            </div>
          `;
          } else {
            html +=`
            <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12 style="text-align:center" align="center">
              <img src="${item.ThirdPartyAppImageLink.Url}" style="width:32px; height:32px;">
            </div>
          `;
          }

          html+=`</div>
          <div class="ms-Grid-row">
            <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12" style="text-align:center">
              <span class="ms-font-m ms-fontColor-neutralPrimary">${item.Title}</span>
            </div>
          </div>
          </a>
        </div>
      `;


    });

    const listContainer: Element = this.domElement.querySelector('#spVideoListContainer');
    listContainer.innerHTML = html;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div style="border: 1px solid ${this.properties.theme}" class="ms-bgColor-neutralLighterAlt" >
      <div>
        <div style="padding:10px; background-color: ${this.properties.theme}; color: ${this.properties.headerfontcolor}" class="ms-font-xl">
          ${escape(this.properties.description)}
        </div>
      </div>

      <div class="ms-Grid" style="padding:10px">
        <div class="ms-Grid-row">
          <div id="spVideoListContainer" class="ms-Grid-col ms-lg12">
          </div>

        </div>
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
