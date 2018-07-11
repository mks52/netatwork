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

import styles from './VideosWebPart.module.scss';
import * as strings from 'VideosWebPartStrings';

import MockHttpClient from './MockHttpClient';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment, EnvironmentType} from '@microsoft/sp-core-library';



export interface IVideosWebPartProps {
  description: string;
  channel: string;
  videoChannel: string;
  videoNumber: number;
  theme: string;
  fontface:string;
  headerfontcolor:string;
}


export interface VideosLists {
  value: VideosList[];
}

export interface VideosList {
  ChannelId: string,
  Description: string,
  DisplayFormUrl: string,
  FileName: string,
  ThumbnailUrl: string
}


export default class VideosWebPartWebPart extends BaseClientSideWebPart<IVideosWebPartProps> {
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
    var url = `https://netatwork212.sharepoint.com/portals/hub/_api/VideoService/Channels`;

    return this.fetchLists(url).then((response) => {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        response.value.map((list: IODataList) => {
            console.log("Found list with title = " + list.Title);
            options.push( { key: list.Id, text: list.Title });
        });

        return options;
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

  private _getListData(): Promise<VideosLists> {
    var url = `https://netatwork212.sharepoint.com/portals/hub/_api/VideoService/Channels('` + this.properties.videoChannel + `')/Videos`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getMockListData(): Promise<VideosLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
      .then((data:VideosList[]) =>{
        var listData: VideosLists = {value: data};
        return listData;
      }) as Promise<VideosLists>;
  }

  private _renderList(items: VideosList[]): void {
    let html1: string ='';
    let html2: string='';
    let videoCount: number= 1;

    html1 += `
    <div style="padding-bottom: 10px;" class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
      <div class="ms-Grid-row">
        <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <a href= "${items[0].DisplayFormUrl}" target="_blank">
            <img src = "${items[0].ThumbnailUrl}" style="width:100%;">
          </a>
        </div>
      </div>

      <div class="ms-Grid-row">
        <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <span class="ms-font-xl" style="color: ${this.properties.fontface}" >${items[0].FileName}</span>
        </div>
      </div>

      <div class="ms-Grid-row">
        <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <p class="ms-font-mPlus">${items[0].Description}</p>
        </div>
      </div>
    </div>`
    items.shift();

    items.forEach((item: VideosList) => {
      if(videoCount<(this.properties.videoNumber)) {
        html2 += `
        <div style="padding-bottom: 10px;" class="ms-Grid-col ms-sm12 ms-md6 ms-lg4">
          <div class="ms-Grid-row">
            <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              <a href= "${item.DisplayFormUrl}" target="_blank">
                <img src = "${item.ThumbnailUrl}" style="width:100%;">
              </a>
            </div>
          </div>

          <div class="ms-Grid-row">
            <div class="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              <span class="ms-font-m" style="color: ${this.properties.fontface}" > ${item.FileName}</span>
            </div>
          </div>
        </div>`
        videoCount++;
      }
    });



    const listContainer: Element = this.domElement.querySelector('#spVideoListContainer1');
    listContainer.innerHTML = html1;
    const listContainer2: Element = this.domElement.querySelector('#spVideoListContainer2');
    listContainer2.innerHTML = html2;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div style="border: 1px solid ${this.properties.theme}" class="ms-bgColor-neutralLighterAlt">
      <div>
        <div style="padding:10px; background-color: ${this.properties.theme}; color: ${this.properties.headerfontcolor}" class="ms-font-xl">
          ${escape(this.properties.description)}
        </div>
      </div>

      <div class="ms-Grid" style="padding:10px">
        <div class="ms-Grid-row">
          <div id="spVideoListContainer1" class="ms-Grid-col ms-lg12">
            <p align="center"><br/><br/><h3>Please edit the web part property and select the appropriate channel to display the videos.</h3><br/><br/></p>
          </div>
          <div id="spVideoListContainer2" class="ms-Grid-col ms-lg12">
          </div>
        </div>
      </div>
    </div>`

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
              groupName: "Video Setting",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Web Part Title"
                }),
                PropertyPaneDropdown('videoChannel', {
                  label: 'Select Channel',
                  options: this.dropdownOptions
                }),
                PropertyPaneDropdown('videoNumber', {
                  label: "Number of Videos",
                  options: [
                    {key: 1, text: '1'},
                    {key: 2, text: '2'},
                    {key: 3, text: '3'},
                    {key: 4, text: '4'},
                    {key: 5, text: '5'},
                    {key: 6, text: '6'},
                    {key: 7, text: '7'},
                  ],
                  selectedKey: 1
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
