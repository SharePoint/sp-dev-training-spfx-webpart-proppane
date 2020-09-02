// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
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
  expansionOptions: any[];
}

export default class HelloPnPControlsWebPart extends BaseClientSideWebPart<IHelloPnPControlsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloPnPControls}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
              <div class="selectedPeople"></div>
              <div class="expansionOptions"></div>
            </div>
          </div>
        </div>
      </div>`;

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
        expansionOptions = expansionOptions + `<li>${option['Region']}: ${option['Comment']} </li>`;
      });
      if (expansionOptions.length > 0) {
        this.domElement.getElementsByClassName('expansionOptions')[0].innerHTML = `<ul>${expansionOptions}</ul>`;
      }
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
                  context: this.context,
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
