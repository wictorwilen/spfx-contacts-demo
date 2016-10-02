'use strict';
import {IContact} from './IContact';
import {Contact} from './Contact';
import {IContactsDataService} from './IContactsDataService';
import { Log } from '@microsoft/sp-client-base';


export class ContactsMockService implements IContactsDataService {
    constructor() {
        Log.info((<any>this["constructor"]).name, 'Initialized');
    }
    public verifyList(listName: string): Promise<boolean> {
        return new Promise<boolean>(resolve => { resolve(true); });
    }
    public getListNames(): Promise<string[]> {
        return new Promise<string[]>(resolve => {
            setTimeout(() => {
                resolve(['Contacts', 'Another list']);
            }, 0);
        });
    }
    public loadData(count: number, listName: string): Promise<IContact[]> {
        return new Promise<IContact[]>((resolve => {
            setTimeout(() => {
                var arr: IContact[] = [];
                for (var i: number = 0; i < count; i++) {
                    arr.push(new Contact('Wictor', 'Wilén', "Architect", 'https://www.geek.com/wp-content/uploads/2010/10/bill_silhouette.jpg'));
                }
                Log.info((<any>this["constructor"]).name, `Loaded ${count} items`);
                resolve(arr);
            }, 0);
        }));
    }
}