// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  IPropertyFieldGroupOrPerson,
  PropertyFieldPeoplePicker,
  PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import styles from './HelloPnPControlsWebPart.module.scss';
import * as strings from 'HelloPnPControlsWebPartStrings';

export interface IHelloPnPControlsWebPartProps {
  description: string;
  people: IPropertyFieldGroupOrPerson[];
  expansionOptions: any[]; // eslint-disable-line @typescript-eslint/no-explicit-any
}

export default class HelloPnPControlsWebPart extends BaseClientSideWebPart<IHelloPnPControlsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloPnPControls} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div class="selectedPeople"></div>
      <div class="expansionOptions"></div>
    </section>`;

    if (this.properties.people && this.properties.people.length > 0) {
      let peopleList: string = '';
      this.properties.people.forEach((person) => {
        peopleList = peopleList + `<li>${person.fullName} (${person.email})</li>`;
      });

      this.domElement.getElementsByClassName('selectedPeople')[0].innerHTML = `<ul>${peopleList}</ul>`;
    }

    if (this.properties.expansionOptions && this.properties.expansionOptions.length > 0) {
      let expansionOptions: string = '';
      this.properties.expansionOptions.forEach((option) => {
        expansionOptions = expansionOptions + `<li>${option.Region}: ${option.Comment} </li>`;
      });
      if (expansionOptions.length > 0) {
        this.domElement.getElementsByClassName('expansionOptions')[0].innerHTML = `<ul>${expansionOptions}</ul>`;
      }
    }
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldPeoplePicker('people', {
                  label: 'Property Pane Field People Picker PnP Reusable Control',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any, // eslint-disable-line @typescript-eslint/no-explicit-any
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),
                PropertyFieldCollectionData('expansionOptions', {
                  key: 'collectionData',
                  label: 'Possible expansion options',
                  panelHeader: 'Possible expansion options',
                  manageBtnLabel: 'Manage expansion options',
                  value: this.properties.expansionOptions,
                  fields: [
                    {
                      id: 'Region',
                      title: 'Region',
                      required: true,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        { key: 'Northeast', text: 'Northeast' },
                        { key: 'Northwest', text: 'Northwest' },
                        { key: 'Southeast', text: 'Southeast' },
                        { key: 'Southwest', text: 'Southwest' },
                        { key: 'North', text: 'North' },
                        { key: 'South', text: 'South' }
                      ]
                    },
                    {
                      id: 'Comment',
                      title: 'Comment',
                      type: CustomCollectionFieldType.string
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
