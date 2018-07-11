import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './SocialFeedsWebPart.module.scss';
import * as strings from 'SocialFeedsWebPartStrings';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment, EnvironmentType} from '@microsoft/sp-core-library';
import {HttpClient, HttpClientResponse} from '@microsoft/sp-http'
import * as $ from 'jquery';
import 'jqueryUi';
import {SPComponentLoader} from '@microsoft/sp-loader';
import * as twttr from 'twitter';
import * as cors from 'cors';
//var twttr: any = require('twitter');
//var $: any = require('jquery');
export interface ISocialFeedsWebPartProps {
  twitter: string;
  facebook: string;
  linkedinPage: string;
  linkedinAccessToken: string;
  theme: string;
  fontface: string;
  headerfontcolor: string;
}

export interface ProfileLists {
  squareLogoUrl: string;
}


export interface BlogsLists {
  values: BlogsList[];
}

export interface BlogsList {
  numLikes: number,
  timestamp: number,
  updateComments: {_total: number},
  updateContent: {
    company: {
      id: number,
      name: string
    },
    companyStatusUpdate: {share: {
      comment: string,
      id: string,
      source: {
        serviceProvider: {name: string},
        serviceProviderShareId: string
      },
      timestamp: number,
      visibility: {code: string}
      content: {
        eyebrowUrl: string,
        mediaKey:string,
        shortenedUrl: string,
        submittedImageUrl: string,
        submittedUrl: string,
        thumbnailUrl: string,
        title: string
      }
    }}
  }
}


export default class SocialFeedsWebPartWebPart extends BaseClientSideWebPart<ISocialFeedsWebPartProps> {

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
    var that = this;
    $.ajax({
      url: 'https://api.linkedin.com/v1/companies/'+ that.properties.linkedinPage+'/updates?oauth2_access_token='+ that.properties.linkedinAccessToken +'&format=jsonp',
      dataType: 'JSONP',
      jsonpCallback: 'callback',
      type: 'GET',
      success: function (data1) {
        $.ajax({
          url: 'https://api.linkedin.com/v1/companies/'+ that.properties.linkedinPage +':(id,name,ticker,description,square-logo-url)?oauth2_access_token=' + that.properties.linkedinAccessToken + '&format=jsonp',
          dataType: 'JSONP',
          jsonpCallback: 'callback',
          type: 'GET',
          success: function (data2) {
            that._renderList(data1.values, data2.squareLogoUrl);
          },
          error: function(err2){
            console.log(err2);
          }
        });
      },
      error: function(err1){
        console.log(err1);
      }
    });
  }





  private _renderList(items: BlogsList[],logo): void {
    let html: string ='<br/>';
    let date1: string;
    items.forEach((item: BlogsList) => {
      date1 =new Date((item.timestamp/1000)*1000).toString()
      var res = date1.substring(4, 15);
      html += `
          <div class="panel panel-default">
            <div class="panel-heading" style="background-color:#ffffff">
              <div class="media">
                <div class="media-left">
                  <a href="#">
                    <img src="${logo}" class="media-object">
                  </a>
                </div>
                <div class="media-body">
                  <a href="https://www.linkedin.com/company/${escape(this.properties.linkedinPage)}/" target="_blank" class="h4 media-heading">${item.updateContent.company.name}</a>
                  <br/>
                  Posted on: ${res}
                </div>
              </div>
            </div>

            <div class="panel-body">`;
              //For Posts (written content)
              if(item.updateContent.companyStatusUpdate.share){
                html+=`<h4>${item.updateContent.companyStatusUpdate.share.comment}</h4>`;
              }

              //For Articles
              let share = item.updateContent.companyStatusUpdate.share;
              if(share.hasOwnProperty("content") && share.content.hasOwnProperty("submittedImageUrl")){
                html+=`
                <div class="list-group">
                  <a class="list-group-item" href="${item.updateContent.companyStatusUpdate.share.content.submittedUrl}">
                    <img src="${item.updateContent.companyStatusUpdate.share.content.submittedImageUrl}" class="img-responsive">
                  </a>
                  <a class="list-group-item" href="${item.updateContent.companyStatusUpdate.share.content.submittedUrl}">
                    <h4 class="text-primary">${item.updateContent.companyStatusUpdate.share.content.title}</h4>
                  </a>
                </div>
                `
              }
              //For Images
              else if(item.updateContent.companyStatusUpdate.share.content){
                html+=`
                  <img src="${item.updateContent.companyStatusUpdate.share.content.eyebrowUrl}" class="img-responsive">`
              }

      html += `
      </div>
      <div class="panel-footer" style="background-color:#ffffff">
        <i class="fa fa-thumbs-o-up" aria-hidden="true"></i> ${item.numLikes}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  <i class="fa fa-comment-o" aria-hidden="true"></i> ${item.updateComments._total}
      </div>

      </div>`;


    });

    const listContainer: Element = this.domElement.querySelector('#tabs3');
    listContainer.innerHTML = html;
  }

  public render(): void {


    this.domElement.innerHTML = `

    <div class="ms-Grid">
      <div class="ms-Grid-row">
        <div class="ms-Grid-col ms-sm12">
          <ul class="nav nav-tabs nav-justified">
            <li class="active"><a href="#tabs-1" data-toggle="tab" ><i class="fa fa-twitter fa-2x txt-primary" aria-hidden="true"></i></a></li>
            <li><a href="#tabs-2" data-toggle="tab"><i class="fa fa-facebook fa-2x" aria-hidden="true"></i></a></li>
            <li><a href="#tabs-3" data-toggle="tab"><i class="fa fa-linkedin fa-2x" aria-hidden="true"></i></a></li>
          </ul>

          <div class="tab-content clearfix" style="border: 1px solid ${this.properties.theme}" border-width: 0 1px 1px; padding: 1px;">
            <div class="tab-pane active" id="tabs-1">
              <a class='twitter-timeline' href='https://twitter.com/${escape(this.properties.twitter)}?ref_src=twsrc%5Etfw' data-tweet-limit="5">&nbsp;</a>
            </div>
            <div class="tab-pane" id="tabs-2">
              <iframe src="https://www.facebook.com/plugins/page.php?href=https%3A%2F%2Fwww.facebook.com%2F${escape(this.properties.facebook)}&tabs=timeline&width=345&height=800&small_header=true&adapt_container_width=true&hide_cover=true&show_facepile=true&appId" width="500" height="800" style="border:none;overflow:hidden" scrolling="no" frameborder="0" allowTransparency="true"></iframe>
            </div>
            <div class="tab-pane" id="tabs-3">
              <div class="col-sm-12" style="padding:10px; background-color: ${this.properties.theme}; color: ${this.properties.headerfontcolor}" class="ms-font-xl">
                 </div>
            </div>
          </div>
        </div>
      </div>
    </div>`;

    twttr.widgets.load();
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('twitter', {
                  label: "Twitter Account Name"
                }),
                PropertyPaneTextField('facebook', {
                  label: 'Facebook Page Name'
                }),
                PropertyPaneTextField('linkedinPage', {
                  label: 'LinkedIn Page (e.g. 12345678)'
                }),
                PropertyPaneTextField('linkedinAccessToken', {
                  label: 'LinkedIn Access Token for Page'
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
            ]
            },
          ]
        }
      ]
    };
  }
}
