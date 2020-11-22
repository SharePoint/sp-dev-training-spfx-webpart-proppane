// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloPropertyPaneWebPart.module.scss';
import * as strings from 'HelloPropertyPaneWebPartStrings';

export interface IHelloPropertyPaneWebPartProps {
  description: string;
  myContinent: string;
  numContinentsVisited: number;
}

export default class HelloPropertyPaneWebPart extends BaseClientSideWebPart<IHelloPropertyPaneWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloPropertyPane}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <p class="${ styles.description}">Continent where I reside: ${escape(this.properties.myContinent)}</p>
              <p class="${ styles.description}">Number of continents I've visited: ${this.properties.numContinentsVisited}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private validateContinents(textboxValue: string): string {
    const validContinentOptions: string[] = ['africa', 'antarctica', 'asia', 'australia', 'europe', 'north america', 'south america'];
    const inputToValidate: string = textboxValue.toLowerCase();

    return (validContinentOptions.indexOf(inputToValidate) === -1)
      ? 'Invalid continent entry; valid options are "Africa", "Antarctica", "Asia", "Australia", "Europe", "North America", and "South America"'
      : '';
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
                PropertyPaneTextField('myContinent', {
                  label: 'Continent where I currently reside',
                  onGetErrorMessage: this.validateContinents.bind(this)
                }),
                PropertyPaneSlider('numContinentsVisited', {
                  label: 'Number of continents I\'ve visited',
                  min: 1, max: 7, showValue: true,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
