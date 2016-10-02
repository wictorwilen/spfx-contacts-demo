import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-client-preview';

import {
  EnvironmentType
} from '@microsoft/sp-client-base';


//import styles from './ReactContacts.module.scss';
import * as strings from 'contactsStrings';
import { IContactsWebPartProps } from '../contacts/IContactsWebPartProps';
import ReactContacts, { IReactContactsProps } from './components/ReactContacts';

import {IContactsDataService} from '../../shared/IContactsDataService';
import {ContactsMockService} from '../../shared/ContactsMockService';
import {ContactsDataService} from '../../shared/ContactsDataService';

export default class ReactContactsWebPart extends BaseClientSideWebPart<IContactsWebPartProps> {

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
    const element: React.ReactElement<IReactContactsProps> = React.createElement(ReactContacts, {
      title: this.properties.title,
      listName: this.properties.listName,
      numberOfContacts: this.properties.numberOfContacts,
      sortAscending: this.properties.sortAscending,
      large: this.properties.large,
      dataService: this._dataService
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {

      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneSlider('numberOfContacts', {
                  min: 0,
                  max: 30,
                  label: strings.NumberFieldLabel
                })
              ]
            },
            {
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneToggle('large', {
                  label: strings.LargeFieldLabel
                }),
                PropertyPaneToggle('sortAscending', {
                  label: strings.SortFieldLabel
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneDescription2
          },
          groups: [
            {
              groupName: strings.ListGroupsName,
              groupFields: [
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
                })
              ]
            }
          ]
        }
      ]
    };

  }
}
