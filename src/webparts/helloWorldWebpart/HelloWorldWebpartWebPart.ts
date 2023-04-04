import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebpartWebPart.module.scss';
import * as strings from 'HelloWorldWebpartWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

export interface IHelloWorldWebpartWebPartProps {
  description: string;
  test: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface TermList {
  value : Terms[];
}

export interface Terms {
  Label : string;
}

export default class HelloWorldWebpartWebPart extends BaseClientSideWebPart<IHelloWorldWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
      <ul class="${styles.list}">
        <li class="${styles.listItem}">
          <span class="ms-font-l">${item.Title}</span>
        </li>
      </ul>`;
    });
  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      });
  }

  private getTerms(): Promise<TermList>{
    const url: string = this.context.pageContext.web.absoluteUrl+'/_api/v2.1/termStore/groups/baf98bb4-6102-4d6b-a837-92d4e72636e4/sets/f6c88c73-1bc1-4019-973f-b034ea41e08a/terms/2e21f62b-594b-4a88-aa9f-a1b6aa7e1f62/children?select=id,labels';
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        console.log('response');
        console.log(response);
        return response.json();
      } 
    }); 
  }

/*
$.ajax({
    url: "https://<your-tenant>.sharepoint.com/_vti_bin/TaxonomyInternalService.json/GetChildTermsInTermSetWithPaging",
    type: "POST",
    headers: {
        "accept": "application/json;odata.metadata=minimal",
        "content-type": "application/json;charset=utf-8",
        "odata-version": "4.0",
        "X-RequestDigest": $("#__REQUESTDIGEST").val()
    },
    data: JSON.stringify({
        lcid: 1033,
        sspId: "guid-here", //Term store ID
        guid: "guid-here", //Term set ID
        includeDeprecated: false,
        pageLimit: 1000,
        pagingForward: false,
        includeCurrentChild: false,
        currentChildId: "00000000–0000–0000–0000–000000000000",
        webId: "00000000–0000–0000–0000–000000000000",
        listId: "00000000–0000–0000–0000–000000000000"
    }),
    success: function (data) {
        console.log("Rejoice, for now you can use REST", data);
    }
})
*/

  private _renderTerms(items: Terms[]): void {
    let html: string = '';
    items.forEach((item: Terms) => {
      html += `
      <ul class="${styles.list}">
        <li class="${styles.listItem}">
          <span class="ms-font-l">label : ${item.Label}</span>
        </li>
      </ul>`;
    });
  
    const listContainer: Element = this.domElement.querySelector('#termsContainer');
    listContainer.innerHTML = html;
  }

  private _renderTermsAsync(): void {
    this.getTerms()
      .then((response) => {
        this._renderTerms(response.value);
      });
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorldWebpart} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <div>Web part description: <strong>${escape(this.properties.description)}</strong></div>
        <div>Web part test: <strong>${escape(this.properties.test)}</strong></div>
        <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
      </div>
      <div id="spListContainer" />
      <div id="termsContainer" />
    </section>`;
    
    this._renderListAsync();
    this._renderTermsAsync();
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
