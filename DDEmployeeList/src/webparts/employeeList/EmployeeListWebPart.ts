import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EmployeeListWebPart.module.scss';
import * as strings from 'EmployeeListWebPartStrings';

import MockHttpClient from './MockHttpClient';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment, EnvironmentType} from '@microsoft/sp-core-library';
import {SPComponentLoader} from '@microsoft/sp-loader';
import { IODataList } from '@microsoft/sp-odata-types';
import { IPropertyPaneData } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPane/IPropertyPane';

export interface IEmployeeListWebPartProps {
  description: string;
  type: string;
  splist: string;
  theme: string;
  fontface: string;
  headerfontcolor: string;
}

export interface EmployeeLists {
  value: EmployeeList[];
}

export interface EmployeeList {
  FirstName: string,
  LastName: string,
  Picture: {
    Url: string
  },
  BirthMonth: number,
  BirthDay: number,
  JoiningDay:  number,
  JoiningMonth: number,
  LastWorkingDay: Date,
  Department: string,
  WorkCity: string,
  WorkState: string
}


export default class EmployeeListWebPartWebPart extends BaseClientSideWebPart<IEmployeeListWebPartProps> {

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
    if(Environment.type === EnvironmentType.SharePoint || EnvironmentType.ClassicSharePoint) {
      this._getListData().then((response) =>
        this._renderList(response.value)
      );
    }
  }

  private _getListData(): Promise<EmployeeLists> {
    switch(this.properties.type) {
      case "birthday":
        var url = `https://netatwork212.sharepoint.com/sites/sandbox/_api/web/Lists/${this.properties.splist}/Items?$orderby=BirthMonth,BirthDay`;
        break;
      case "anniversary":
        var url = `https://netatwork212.sharepoint.com/sites/sandbox/_api/web/Lists/${this.properties.splist}/Items?$orderby=JoiningMonth,JoiningDay`;
        break;
      case "departing":
        var url = `https://netatwork212.sharepoint.com/sites/sandbox/_api/web/Lists/${this.properties.splist}/Items?$orderby=BirthMonth,BirthDay`;
        break;
    }


    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log(response.json.toString);
        return response.json();
      });
  }


  private _getMockListData(): Promise<EmployeeLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
      .then((data:EmployeeList[]) =>{
        var listData: EmployeeLists = {value: data};
        return listData;
      }) as Promise<EmployeeLists>;
  }

  private _renderList(items: EmployeeList[]): void {
    let html: string ='';
    let months = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    let cdate = new Date();
    let thisMon = cdate.getMonth() + 1;
    let thisDay = cdate.getDate();
    // console.log("Current Date = " + thisMon + "/" + thisDay);
    let status = false;
    let i = 0;
    let start = 0;
    let x = 0;
    if(this.properties.type == "birthday") {

      items.forEach((item: EmployeeList) => {
        let bday = item.BirthDay;
        let bmon = months[item.BirthMonth];
        let bmonNum = item.BirthMonth;
        if(!status) {
          if(bmonNum == thisMon) {
            if(bday >= thisDay) {
              start = i;
              status = true;
            }
          } else if(bmonNum > thisMon) {
            start = i;
            status = true;
          }
        }
        i++;
      });

      let show = 1;
      let sameday:boolean = true;
      let item:EmployeeList = items[start];
      x = Math.min(items.length, 5);
      while((show <= x || sameday) && show <= items.length) {
        let bday = item.BirthDay;
        let bmon = months[item.BirthMonth];
        let bmonNum = item.BirthMonth;
        let display:boolean = true;


          html += `
          <div class="col-sm-12" style="padding-bottom:10px">
            <div class="media">
              <div class="media-left">
                <img class="media-object" src="${item.Picture.Url}" style="width:96px; height:96px;">
              </div>
              <div class="media-body" style="vertical-align:top; padding-top:10px;">
                <h3 class="media-heading" style="color: ${this.properties.fontface}">${item.FirstName} ${item.LastName} <span class="badge">${bday} ${bmon}</span></h3>`
                if(item.Department != null) {
                  html+=`<h5>${item.Department}</h5>`;
                }
                if(item.WorkCity != null && item.WorkState != null) {
                  html+=`<h5>${item.WorkCity}, ${item.WorkState}</h5>`;
                }
                html += `
              </div>
            </div>
          </div>
          `;
          // console.log(show + "," + start + "," + sameday);
          show++;


        start++;

        if (start > items.length - 1) {
          start = 0;
        }
        item = items[start];
        sameday = (thisMon == item.BirthMonth && thisDay == item.BirthDay);
      }
    } else if (this.properties.type == "anniversary") {
      items.forEach((item: EmployeeList) => {
        let aday = item.JoiningDay;
        let amon = months[item.JoiningMonth];
        let amonNum = item.JoiningMonth;
        if(!status) {
          if(amonNum == thisMon) {
            if(aday >= thisDay) {
              start = i;
              status = true;
            }
          } else if(amonNum > thisMon) {
            start = i;
            status = true;
          }
        }
        i++;
      });

      let show = 1;
      let sameday:boolean = true;
      let item:EmployeeList = items[start];
      x = Math.min(items.length, 5);
      while((show <= x || sameday) && show <= items.length) {
        let aday = item.JoiningDay;
        let amon = months[item.JoiningMonth];
        let amonNum = item.JoiningMonth;
        let display:boolean = true;


          html += `
          <div class="col-sm-12" style="padding-bottom:10px">
            <div class="media">
              <div class="media-left">
                <img class="media-object" src="${item.Picture.Url}" style="width:96px; height:96px;">
              </div>
              <div class="media-body" style="vertical-align:top; padding-top:10px;">
                <h3 class="media-heading" style="color: ${this.properties.fontface}">${item.FirstName} ${item.LastName} <span class="badge">${aday} ${amon}</span></h3>`
                if(item.Department != null) {
                  html+=`<h5>${item.Department}</h5>`;
                }
                if(item.WorkCity != null && item.WorkState != null) {
                  html+=`<h5>${item.WorkCity}, ${item.WorkState}</h5>`;
                }
                html += `
              </div>
            </div>
          </div>
          `;
          // console.log(show + "," + start + "," + sameday);
          show++;


        start++;

        if (start > items.length - 1) {
          start = 0;
        }
        item = items[start];
        sameday = (thisMon == item.JoiningMonth && thisDay == item.JoiningDay);
      }
    } else if (this.properties.type == "departing") {

    }



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
    if (!this.listsFetched) {
      this.fetchOptions().then((response) => {
        this.dropdownOptions = response;
        this.listsFetched = true;
        // now refresh the property pane, now that the promise has been resolved..
        this.context.propertyPane.refresh();
      });
    }


    return {
      pages: [
        {

          groups: [
            {
              groupName : "Title",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title of the web part"
                })
              ]
            },
            {
              groupName : "Type of List",
              groupFields: [
                PropertyPaneDropdown('type', {
                  label: "Birthdays or Anniversaries?",
                  options: [
                    {key: 'birthday', text: 'Birthday List'},
                    {key: 'anniversary', text: 'Anniversary List'}
                    // {key: 'departing', text: 'Departing Employees\' List'}
                  ],
                  selectedKey: 'birthday'
                }),
              ]
            },
            {
              groupName : "List",
              groupFields: [
                PropertyPaneDropdown('splist', {
                  label: 'Select the list from your site contents. The list should be created from the template BirthdayAnniversaryTemplate',
                  options: this.dropdownOptions,

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
