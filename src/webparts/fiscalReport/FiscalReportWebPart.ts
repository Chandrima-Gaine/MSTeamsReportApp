import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FiscalReportWebPart.module.scss';
import * as strings from 'FiscalReportWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IFiscalReportWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Expense: number;
  Revenue: number;
  Profit: number;
  YOYGrowth: string;
}

export default class FiscalReportWebPart extends BaseClientSideWebPart<IFiscalReportWebPartProps> {
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('CompanyFinancialReport')/Items",SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
        return response.json();
        });
    }
    private _renderListAsync(): void {
    
      if (Environment.type == EnvironmentType.SharePoint || 
               Environment.type == EnvironmentType.ClassicSharePoint) {
       this._getListData()
         .then((response) => {
           this._renderList(response.value);
         });
     }
   }

   private _renderList(items: ISPList[]): void {
    let html: string = '<table border=1 width=100% style="border-collapse: collapse; border-color: white">';
    html += '<th>Quarter</th> <th>Expense($)</th> <th>Revenue($)</th> <th>Profit($)</th> <th>Y-O-Y Growth</th>';
    items.forEach((item: ISPList) => {
      html += `
      <tr style="text-align:  center">            
          <td>${item.Title}</td>
          <td>${item.Expense}</td>
          <td>${item.Revenue}</td>
          <td>${item.Profit}</td>
          <td>${item.YOYGrowth}</td>            
          </tr>
          `;
    });
    html += '</table>';
  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }
    public render(): void {
    let reportTitle: string = '';
    let spTitle: string = '';
    let spsubTitle: string = '';
    let siteTabTitle: string = '';

    if (this.context.sdks.microsoftTeams) 
    {
      // We have teams context for the web part
      reportTitle = "2019 Fiscal Year Report of Company";
      this._renderListAsync();
    }
    else
    {
      // We are rendered in normal SharePoint context
      spTitle = "Welcome to SharePoint!";
      spsubTitle = "We are in the context of following site: " + this.context.pageContext.web.title;
      siteTabTitle = "2019 Fiscal Year Report of Company";
      this._renderListAsync();
    }
        this.domElement.innerHTML = `
          <div class="${ styles.fiscalReport }">
            <div class="${ styles.container }">
              <div class="${ styles.row }">

                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <span class="${ styles.title }">${reportTitle}</span>
                </div>
    
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <span class="${ styles.title }">${spTitle}</span>
                  <p class="${ styles.subTitle }">${spsubTitle}</p>
                  <p class="${ styles.description }">${siteTabTitle}</p>
                </div>
                
                <div class="${ styles.row }">
                  <div id="spListContainer" />
                </div>

              </div>
            </div>
          </div>`;
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
            }
          ]
        }
      ]
    };
  }
}
