// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IContinentSelectorProps } from './components/IContinentSelectorProps';
import ContinentSelector from './components/ContinentSelector';
import {
  IPropertyPaneContinentSelectorProps,
  IPropertyPaneContinentSelectorInternalProps,
} from './';


export class PropertyPaneContinentSelector implements IPropertyPaneField<IPropertyPaneContinentSelectorProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public properties: IPropertyPaneContinentSelectorInternalProps;
  private element: HTMLElement;

  constructor(public targetProperty: string, properties: IPropertyPaneContinentSelectorProps) {
    this.properties = {
      key: properties.label,
      label: properties.label,
      disabled: properties.disabled,
      selectedKey: properties.selectedKey,
      onPropertyChange: properties.onPropertyChange,
      onRender: this.onRender.bind(this)
    };
  }

  public render(): void {
    if (!this.element) {
      return;
    }
  }

  private onRender(element: HTMLElement): void {
    if (!this.element) {
      this.element = element;
    }

    const reactElement: React.ReactElement<IContinentSelectorProps> = React.createElement(ContinentSelector, <IContinentSelectorProps>{
      label: this.properties.label,
      onChanged: this.onChanged.bind(this),
      selectedKey: this.properties.selectedKey,
      disabled: this.properties.disabled,
      stateKey: new Date().toString() // hack to allow for externally triggered re-rendering
    });
    ReactDom.render(reactElement, element);
  }

  private onChanged(option: IDropdownOption, index?: number): void {
    this.properties.onPropertyChange(this.targetProperty, option.key);
  }
}
