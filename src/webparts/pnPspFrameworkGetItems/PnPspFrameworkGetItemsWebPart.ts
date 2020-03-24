import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnPspFrameworkGetItemsWebPart.module.scss';
import pnp from 'sp-pnp-js';
import * as strings from 'PnPspFrameworkGetItemsWebPartStrings';
import MockHttpClient from './MockHttpClient';

export interface IPnPspFrameworkGetItemsWebPartProps {
  description: string;
}
export interface ISPList {
  Title: string;
  EmployeeName: string;
  Experience: number
  Branch: string;
}


export default class PnPspFrameworkGetItemsWebPart extends BaseClientSideWebPart <IPnPspFrameworkGetItemsWebPartProps> {


  


  private _getListData(): Promise<ISPList[]> {
    return pnp.sp.web.lists.getByTitle("EmployeeDetails").items.get().then((response) => {
     
       return response;
     });
       
    }

    private _renderListAsync(): void {
       
      
       
         this._getListData()
         .then((response:any) => {
           this._renderList(response);
         });
    }   


   private _renderList(items: ISPList[]): void {
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
    html += `<th>EmployeeId</th><th>EmployeeName</th><th>Experience</th><th>Branch</th>`;
    items.forEach((item: ISPList) => {
      html += `
           <tr>
          <td>${item.Title}</td>
          <td>${item.EmployeeName}</td>
          <td>${item.Experience}</td>
           <td>${item.Branch}</td>
           </tr>
           `;
     });
     html += `</table>`;
     const listContainer: Element = this.domElement.querySelector('#spListContainer');
     listContainer.innerHTML = html;
    }
    

  public render(): void {
    this.domElement.innerHTML = `  
    <div class="${styles.pnPspFrameworkGetItems}">  
 <div class="${styles.container}">  
   <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
     <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
       <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Welcome to SharePoint Framework Development</span>  
         
       <p class="ms-font-l ms-fontColor-white" style="text-align: center">Demo : Retrieve Employee Data from SharePoint List</p>  
     </div>  
   </div>  
   <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
   <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Employee Details</div>  
   <br>  
<div id="spListContainer" />  
   </div>  
 </div>  
</div>`;  
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
          }
        ]
      }
    ]
  };
}
}
