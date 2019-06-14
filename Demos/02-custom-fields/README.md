# DEMO: Building custom property pane fields

In this demo, you will extend the property pane by creating your own custom field control.

### Create custom dropdown control of continents

1. Open a command prompt and change to the root folder of the project created in the last demo.

    > NOTE: If you did not create the project in the previous demo, you will find a complete working version that you can start from in the **[Demos\01-proppane](../01-proppane)** folder.

1. Within the project, locate the **src** folder and create a subfolder **controls**.
1. Create a new folder **PropertyPaneContinentSelector** within the **controls** folder to contain the new drop down control.
1. Custom property pane controls should be created using React & Fabric React to match the same user interface as the native controls. A custom React component is included in the **LabFiles** associated with this lab.

    Copy the folder **[../../LabFiles/components](../../LabFiles/components)** into the **src/controls/PropertyPaneContinentSelector** folder in the project.

1. Create an interface for the public properties that can be set as options on the custom property pane control you will build in a moment:
    1. Create a new file **IPropertyPaneContinentSelectorProps.ts** in the **src/controls/PropertyPaneContinentSelector** folder.
    1. Add the following code to the **IPropertyPaneContinentSelectorProps.ts** file:

        ```ts
        import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

        export interface IPropertyPaneContinentSelectorProps {
          label: string;
          onPropertyChange: (propertyPath: string, newValue: any) => void;
          selectedKey: string | number;
          disabled: boolean;
        }
        ```

1. Create another interface that merges both the properties for the custom control and the custom property pane control interface:
    1. Create a new file **IPropertyPaneContinentSelectorInternalProps.ts** in the **src/controls/PropertyPaneContinentSelector** folder.
    1. Add the following code to the **IPropertyPaneContinentSelectorInternalProps.ts** file:

        ```ts
        import { IPropertyPaneCustomFieldProps } from '@microsoft/sp-webpart-base';
        import { IPropertyPaneContinentSelectorProps } from './IPropertyPaneContinentSelectorProps';

        export interface IPropertyPaneContinentSelectorInternalProps extends IPropertyPaneCustomFieldProps, IPropertyPaneContinentSelectorProps { }
        ```

1. Simplify importing the public parts of the control by creating a barrel that exports everything:
    1. Create a new file **index.ts** in the **src/controls/PropertyPaneContinentSelector** folder.
    1. Add the following code to the file:

        ```ts
        export * from './IPropertyPaneContinentSelectorProps';
        export * from './IPropertyPaneContinentSelectorInternalProps';
        export * from './PropertyPaneContentSelector';
        ```

1. Now create the custom property pane control. This control will load a React control and wire up the properties provided to the control to the public properties on the React control:
    1. Create a new file named **PropertyPaneContentSelector.ts** in the **src/controls/PropertyPaneContinentSelector** folder.
    1. Add the following `import` statements to the top of the file:

        ```ts
        import * as React from 'react';
        import * as ReactDom from 'react-dom';
        import { 
          IPropertyPaneField, 
          PropertyPaneFieldType 
        } from '@microsoft/sp-webpart-base';
        import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
        import { IContinentSelectorProps } from './components/IContinentSelectorProps';
        import ContinentSelector from './components/ContinentSelector';
        import { 
          IPropertyPaneContinentSelectorProps,
          IPropertyPaneContinentSelectorInternalProps,
        } from './';
        ```

    1. Next, declare the new class that implements the `IPropertyPaneField` interface provided by the SPFx API with a few class members:

        ```ts
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
        }
        ```

    1. Add the following two methods to the `PropertyPaneContinentSelector` class that will create the React component and render it on the page as well as wire up the change event when the selection changes:

        ```ts
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
        ```

### Implement the custom property pane control

With the custom property pane control created, you can now replace the existing text box control with the new control.

1. Locate and open the **src/webparts/helloPropertyPane/HelloPropertyPaneWebPart.ts** file.
1. Add the following `import` statement to the top of the file after the existing `import` statements:

    ```ts
    import {
      PropertyPaneContinentSelector,
      IPropertyPaneContinentSelectorProps
    } from '../../controls/PropertyPaneContinentSelector';
    ```

1. Locate the `getPropertyPaneConfiguration()` method in the web part, then find the existing `PropertyPaneTextField` that's bound to the **myContinent** property. Comment this control out
1. Add the following custom control to the property pane:

    ```ts
    new PropertyPaneContinentSelector('myContinent', <IPropertyPaneContinentSelectorProps>{
      label: 'Continent where I currently reside',
      disabled: false,
      selectedKey: this.properties.myContinent,
      onPropertyChange: this.onContinentSelectionChange.bind(this),
    }),
    ```

1. Add the following method to the `HelloPropertyPaneWebPart` class to handle the event when a user changes the selection in the control. This will update the property on the web part:

    ```ts
    private onContinentSelectionChange(propertyPath: string, newValue: any): void {
      const oldValue: any = this.properties[propertyPath];
      this.properties[propertyPath] = newValue;
      this.render();
    }
    ```

1. Now test the web part by executing `gulp serve` (*if the local web server is not already running*). You will see the new control selector and notice the values change in the web part when you change the selection.

    ![Screenshot of custom property pane field control selector](../../Images/EditPropPane-CustomControl-01.png)