declare interface IMyProfileWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'MyProfileWebPartStrings' {
  const strings: IMyProfileWebPartStrings;
  export = strings;
}
