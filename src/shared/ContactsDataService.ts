'use strict';
import {IContact} from './IContact';
import {Contact} from './Contact';
import {IContactsDataService} from './IContactsDataService';

import { HttpClient } from '@microsoft/sp-client-base';
import { IWebPartContext } from '@microsoft/sp-client-preview';

import { Log } from '@microsoft/sp-client-base';


export class ContactsDataService implements IContactsDataService {
    private _httpClient: HttpClient;
    private _webAbsoluteUrl: string;

    public constructor(webPartContext: IWebPartContext) {
        this._httpClient = webPartContext.httpClient as any; // tslint:disable-line:no-any
        this._webAbsoluteUrl = webPartContext.pageContext.web.absoluteUrl;
        Log.info(this["constructor"].toString(), 'Initialized');
    }

    public verifyList(listName: string): Promise<boolean> {
        return this._httpClient.get(this._webAbsoluteUrl + `/_api/Lists/GetByTitle('${listName}')`)
            .then((response: Response) => {
                return response.status === 200;
            }).catch(error => {
                return false;
            });
    }

    public getListNames(): Promise<string[]> {
        return this._httpClient.get(this._webAbsoluteUrl + `/_api/Lists/?$select=Title`)
            .then((response: Response) => {
                var arr: string[] = [];
                return response.json().then((data) => {
                    data.value.forEach(l => {
                        arr.push(l.Title);
                    });
                    return arr;
                });
            });
    }


    public loadData(count: number, listName: string): Promise<IContact[]> {
        return this._httpClient.get(this._webAbsoluteUrl + `/_api/Lists/GetByTitle('${listName}')/Items?$top=${count}`)
            .then((response: Response) => {

                var arr: IContact[] = [];
                return response.json().then((data) => {
                    data.value.forEach(c => {
                        arr.push(
                            new Contact(
                                c.First_x0020_name,
                                c.Title,
                                c.Job_x0020_Title,
                                c.Image.Url)
                        );
                    });
                    Log.info((<any>this["constructor"]).name, `Loaded ${data.value.length} items`);
                    return arr;
                });
            });
    }
}