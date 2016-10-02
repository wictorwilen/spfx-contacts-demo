'use strict';
import {IContact} from './IContact';

export class Contact implements IContact {
    public initials: string;

    constructor(
        public firstName: string,
        public lastName: string,
        public jobTitle: string,
        public photoUrl: string
    ) {
        this.initials = firstName[0] + lastName[0];
    }
}