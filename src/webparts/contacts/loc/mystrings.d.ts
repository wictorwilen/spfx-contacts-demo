declare interface IContactsStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListNameFieldLabel: string;
  NumberFieldLabel: string;
  SortFieldLabel: string;
  LargeFieldLabel: string;
  TitleFieldLabel: string;
  Loading: string;
  PropertyPaneDescription2: string;
  ListGroupsName: string;
  LayoutGroupName: string;
}

declare module 'contactsStrings' {
  const strings: IContactsStrings;
  export = strings;
}
