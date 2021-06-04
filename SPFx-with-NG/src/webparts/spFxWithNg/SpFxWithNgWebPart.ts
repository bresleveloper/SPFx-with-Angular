import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxWithNgWebPart.module.scss';
import * as strings from 'SpFxWithNgWebPartStrings';


import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions  } from '@microsoft/sp-http';


/****  ANGULAR  *****/
require('./ng files/runtime.js')
require('./ng files/polyfills.js')
require('./ng files/styles.js')
require('./ng files/vendor.js')
require('./ng files/main.js')




export interface ISpFxWithNgWebPartProps {
  description: string;
}

export default class SpFxWithNgWebPart extends BaseClientSideWebPart<ISpFxWithNgWebPartProps> {

  public render(): void {
    window['ctx'] = this.context
    window['spPost'] = this.spPost.bind(this)
    this.domElement.innerHTML = `<app-any-name></app-any-name>`
  }

  spPost(url:string, payload:object):Promise<any>{
    return new Promise((resolve, reject) => {
      const spOpts: ISPHttpClientOptions = {
        /*body: JSON.stringify({
          '__metadata': { 'type': 'SP.List' },
          'BaseTemplate': 100,
          'Title': listName
          }),*/
        body : JSON.stringify(payload),
        headers: { 
          'Content-Type': 'application/json;odata=verbose',
          "Accept": "application/json;odata=verbose",
        }
      };

      this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts)
        .then((response: SPHttpClientResponse) => {
          console.log("spHttpClient.post", response.status, response);

          //response.json() returns a promise so you get access to the json in the resolve callback.
          response.json().then((responseJSON: JSON) => {
            console.log(responseJSON);
            resolve(responseJSON);
          });
        })
    })
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
