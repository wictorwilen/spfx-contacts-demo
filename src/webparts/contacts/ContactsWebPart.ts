import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
  , PropertyPaneDropdown
  , IPropertyPaneDropdownOption
} from '@microsoft/sp-client-preview';

import {
  EnvironmentType
} from '@microsoft/sp-client-base';

import { Log } from '@microsoft/sp-client-base';

//import styles from './Contacts.module.scss';
import * as strings from 'contactsStrings';
import { IContactsWebPartProps } from './IContactsWebPartProps';

import {IContact} from '../../shared/IContact';
import {IContactsDataService} from '../../shared/IContactsDataService';
import {ContactsMockService} from '../../shared/ContactsMockService';
import {ContactsDataService} from '../../shared/ContactsDataService';

// npm install handlebars --save
// tsd install handlebars --save
import * as Handlebars from 'handlebars';
import * as Templates from './ContactsWebPart.Templates';
//import * as jQuery from 'jquery';

export default class ContactsWebPart extends BaseClientSideWebPart<IContactsWebPartProps> {

  private _dataService: IContactsDataService;

  public constructor(context: IWebPartContext) {
    super(context);

    const isDebug: boolean =
      DEBUG && (this.context.environment.type === EnvironmentType.Test || this.context.environment.type === EnvironmentType.Local);

    this._dataService = isDebug
      ? new ContactsMockService()
      : new ContactsDataService(this.context);
  }

  public render(): void {
    //jQuery(document).ready(()=>{});

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, strings.Loading);

    this._dataService.loadData(this.properties.numberOfContacts, this.properties.listName).then(response => {
      // npm install handlebars-loader --save-dev
      // add the copy-static-assets.json file
      //var hb: any = this.properties.large ? require('handlebars!./templates/large.hbs') : require('handlebars!./templates/small.hbs');

      var hb: any = Handlebars.compile(this.properties.large ? Templates.TemplateLarge : Templates.TemplateSmall);

      this.context.statusRenderer.clearLoadingIndicator(this.domElement);

      this.domElement.innerHTML = `<h1 class='ms-fontSize-xxl'>${this.properties.title == undefined ? 'Contacts' : this.properties.title}</h1>`;

      response.sort(this.properties.sortAscending ? this.sortContact : this.sortContactReverse).forEach(c => {
        this.domElement.innerHTML += hb(c);
      });
    }).catch((error: any) => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.context.statusRenderer.renderError(this.domElement, error);
    });
  }

  protected sortContact(a: IContact, b: IContact): number {
    if (a.lastName > b.lastName) {
      return 1;
    } else if (a.lastName < b.lastName) {
      return -1;
    }
    return 0;
  }
  protected sortContactReverse(a: IContact, b: IContact): number {
    if (a.lastName > b.lastName) {
      return -1;
    } else if (a.lastName < b.lastName) {
      return 1;
    }
    return 0;
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  onGetErrorMessage: (value): Promise<string> => {
                    return new Promise<string>((resolve, reject) => {
                      this._dataService.verifyList(value).then(result => {
                        if (result) {
                          return resolve("");
                        }
                        return resolve("List not found");
                      });
                    });
                  }
                }),
                /*
                                PropertyPaneDropdown('listName', {
                                  label: strings.ListNameFieldLabel,
                                  options: this._listNames
                                }),
                */
                PropertyPaneSlider('numberOfContacts', {
                  min: 0,
                  max: 30,
                  label: strings.NumberFieldLabel
                }),
                PropertyPaneToggle('large', {
                  label: strings.LargeFieldLabel
                }),
                PropertyPaneToggle('sortAscending', {
                  label: strings.SortFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }


  private _listNames: IPropertyPaneDropdownOption[] = [];

  public onInit<T>(): Promise<T> {
    Log.info((<any>this["constructor"]).name, 'IPropertyPaneDropdownOption()');

    return new Promise<T>((resolve: (args: T) => void, reject: (error: Error) => void) => {
      this._dataService.getListNames().then(lists => lists.forEach(list => {
        Log.info((<any>this["constructor"]).name, 'Has data in IPropertyPaneDropdownOption()');
        this._listNames.push(<IPropertyPaneDropdownOption>{
          text: list,
          key: list
        });
        resolve(undefined);
      }))
    });


    /*
        this._dataService.getListNames().then(lists => lists.forEach(list => {
           Log.info((<any>this["constructor"]).name, 'Has data in IPropertyPaneDropdownOption()');
           this._listNames.push(<IPropertyPaneDropdownOption>{
             text: list,
             key: list
           });
         }));
         Log.info((<any>this["constructor"]).name, 'Resolve IPropertyPaneDropdownOption()');

         return Promise.resolve();
         */
  }


  /*
    protected get disableReactivePropertyChanges(): boolean {
      return true;
    }
    */
}
