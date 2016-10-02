'use strict';
import {IContact} from './IContact';

export interface IContactsDataService {
    getListNames(): Promise<string[]>;
    verifyList(listName: string): Promise<boolean>;
    loadData(count: number, listName: string): Promise<IContact[]>;
}