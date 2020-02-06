// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as React from 'react';
import {
  Dropdown,
  IDropdownOption
} from 'office-ui-fabric-react/lib/components/Dropdown';
import { IContinentSelectorProps } from './IContinentSelectorProps';
import { IContinentSelectorState } from './IContinentSelectorState';

export default class ContinentSelector extends React.Component<IContinentSelectorProps, IContinentSelectorState> {
  private selectedKey: React.ReactText;

  constructor(props: IContinentSelectorProps, state: IContinentSelectorState) {
    super(props);
    this.selectedKey = props.selectedKey;
    this.state = { options: [] };
  }

  public componentDidMount(): void {
    this.loadOptions();
  }

  public loadOptions(): void {
    let continents: IDropdownOption[] = [
      { "key": "Africa", "text": "Africa" },
      { "key": "Antarctica", "text": "Antarctica" },
      { "key": "Asia", "text": "Asia" },
      { "key": "Australia", "text": "Australia" },
      { "key": "Europe", "text": "Europe" },
      { "key": "North America", "text": "North America" },
      { "key": "South America", "text": "South America" },
    ];
    this.setState({ options: continents });
  }

  public render(): JSX.Element {
    return (
      <div>
        <Dropdown label={this.props.label}
                  disabled={this.props.disabled}
                  selectedKey={this.selectedKey}
                  options={this.state.options}
                  onChanged={this.onChanged.bind(this)} />
      </div>
    );
  }

  private onChanged(option: IDropdownOption, index?: number): void {
    this.selectedKey = option.key;
    const options: IDropdownOption[] = this.state.options;
    options.forEach((opt: IDropdownOption): void => {
      if (opt.key !== option.key) {
        opt.selected = false;
      }
    });
    this.setState((prevState: IContinentSelectorState, props: IContinentSelectorProps): IContinentSelectorState => {
      prevState.options = options;
      return prevState;
    });
    if (this.props.onChanged) {
      this.props.onChanged(option, index);
    }
  }

}