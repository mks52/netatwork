declare interface IMyAppsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'MyAppsWebPartStrings' {
  const strings: IMyAppsWebPartStrings;
  export = strings;
}
