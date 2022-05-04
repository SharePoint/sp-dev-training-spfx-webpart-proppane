declare interface IHelloPropertyPaneWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'HelloPropertyPaneWebPartStrings' {
  const strings: IHelloPropertyPaneWebPartStrings;
  export = strings;
}
