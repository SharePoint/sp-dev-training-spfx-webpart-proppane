declare interface IHelloPnPControlsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'HelloPnPControlsWebPartStrings' {
  const strings: IHelloPnPControlsWebPartStrings;
  export = strings;
}
