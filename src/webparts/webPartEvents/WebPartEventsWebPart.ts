import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  IWebPartData,
  IHtmlProperties,
  IWebPartEvent
} from '@microsoft/sp-client-preview';
import {
  DisplayMode
} from '@microsoft/sp-client-base';

import { Log } from '@microsoft/sp-client-base';

import * as strings from 'webPartEventsStrings';
import { IWebPartEventsWebPartProps } from './IWebPartEventsWebPartProps';

export default class WebPartEventsWebPart extends BaseClientSideWebPart<IWebPartEventsWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
    Log.info((<any>this["constructor"]).name, 'constructor()');
  }

  public render(): void {
    Log.info((<any>this["constructor"]).name, 'render()');
    this.domElement.innerHTML = `Client-side web parts rocks!`;
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    Log.info((<any>this["constructor"]).name, 'propertyPaneSettings()');
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
                  label: strings.DescriptionFieldLabel,
                  onGetErrorMessage: (value): Promise<string> => {
                    Log.info((<any>this["constructor"]).name, 'onGetErrorMessage()');
                    return new Promise<string>((resolve, reject) => {
                      return resolve("");
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

  protected get disableReactivePropertyChanges(): boolean {
    Log.info((<any>this["constructor"]).name, 'disableReactivePropertyChanges()');
    return false;
  }

  protected deserialize(data: IWebPartData): IWebPartEventsWebPartProps {
    Log.info((<any>this["constructor"]).name, 'deserialize()');
    return super.deserialize(data);
  }

  protected onInit<IWebPartEventsWebPartProps>(): Promise<IWebPartEventsWebPartProps> {
    Log.info((<any>this["constructor"]).name, 'onInit()');
    return super.onInit();
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    Log.info((<any>this["constructor"]).name, `onDisplayModeChanged(${oldDisplayMode})`);
  }

  protected onBeforeSerialize(): IHtmlProperties {
    //Log.verbose((<any>this["constructor"]).name, 'onBeforeSerialize()');
    return super.onBeforeSerialize();
  }

  protected onEvent<IWebPartEventsWebPartProps>(eventName: string, eventObject: IWebPartEvent<IWebPartEventsWebPartProps>): void {
    Log.info((<any>this["constructor"]).name, `onDisplayModeChanged('${eventName}')`);

    return super.onEvent(eventName, eventObject);
  }

  protected dispose(): void {
    Log.info((<any>this["constructor"]).name, `dispose()`);
    super.dispose();
  }

  protected onPropertyChange(propertyPath: string, newValue: any): void {
    Log.info((<any>this["constructor"]).name, `onPropertyChange('${propertyPath}', '${newValue}')`);
    super.onPropertyChange(propertyPath, newValue);
  }

  protected onPropertyPaneConfigurationStart(): void {
    Log.info((<any>this["constructor"]).name, `onPropertyPaneConfigurationStart()`);
    super.onPropertyPaneConfigurationStart();
  }
  protected onPropertyPaneConfigurationComplete(): void {
    Log.info((<any>this["constructor"]).name, `onPropertyPaneConfigurationComplete()`);
    super.onPropertyPaneConfigurationStart();
  }
  protected onAfterPropertyPaneChangesApplied(): void {
    Log.info((<any>this["constructor"]).name, `onAfterPropertyPaneChangesApplied()`);
    super.onPropertyPaneConfigurationStart();
  }
  protected onPropertyPaneRendered(): void {
    Log.info((<any>this["constructor"]).name, `onPropertyPaneRendered()`);
    super.onPropertyPaneConfigurationStart();
  }
}
