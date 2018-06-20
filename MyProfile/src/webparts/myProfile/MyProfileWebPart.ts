import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyProfileWebPart.module.scss';
import * as strings from 'MyProfileWebPartStrings';
import MockHttpClient from './MockHttpClient';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IMyProfileWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  EmployeeId: string;
  EmployeeName: string;
  Experience: string;
  Location: string;
}

export default class MyProfileWebPart extends BaseClientSideWebPart<IMyProfileWebPartProps> {

  private _renderListAsync(): void {

    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }


  private _renderList(items: ISPList[]): void {
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
    html += `<th>EmployeeId</th><th>EmployeeName</th><th>Experience</th><th>Location</th>`;
    items.forEach((item: ISPList) => {
      html += `  
           <tr>  
          <td>${item.EmployeeId}</td>  
          <td>${item.EmployeeName}</td>  
          <td>${item.Experience}</td>  
          <td>${item.Location}</td>  
          </tr>  
          `;
    });
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.myProfile }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint Manikanta!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
        <div id="spListContainer" /> </div>
      </div>`;
      this._renderListAsync();
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {
      const listData: ISPLists = {
        value:
          [
            { EmployeeId: 'E123', EmployeeName: 'John', Experience: 'SharePoint', Location: 'India' },
            { EmployeeId: 'E567', EmployeeName: 'Martin', Experience: '.NET', Location: 'Qatar' },
            { EmployeeId: 'E367', EmployeeName: 'Luke', Experience: 'JAVA', Location: 'UK' }
          ]
      };
      return listData;
    }) as Promise<ISPLists>;
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
