// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface IContinentSelectorProps {
  label: string;
  onChanged: (option: IDropdownOption, index?: number) => void;
  selectedKey: string | number;
  disabled: boolean;
  stateKey: string;
}
