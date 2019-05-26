# DEMO: Working with the web part property pane

In this exercise, you will get hands-on experience in manipulating the property pane for a SPFx client-side web part in working with controls, groups and pages.

## Create a new SPFx Solution & Web Part

> NOTE: The instructions below assume you are using v1.8.2 of the SharePoint Framework Yeoman generator.

1. Open a command prompt and change to the folder where you want to create the project.
1. Run the SharePoint Framework Yeoman generator by executing the following command:

    ```shell
    yo @microsoft/sharepoint
    ```

    Use the following to complete the prompt that is displayed:

    - **What is your solution name?:** HelloPropertyPane
    - **Which baseline packages do you want to target for your component(s)?:** SharePoint Online only (latest)
    - **Where do you want to place the files?:** Use the current folder
    - **Do you want to allow the tenant admin the choice of being able to deploy the solution to all sites immediately without running any feature deployment or adding apps in sites?:** No
    - **Will the components in the solution require permissions to access web APIs that are unique and not shared with other components in the tenant?:** No    
    - **Which type of client-side component to create?:** WebPart
    - **What is your Web part name?:** HelloPropertyPane
    - **What is your Web part description?:** HelloPropertyPane description
    - **Which framework would you like to use?:** No JavaScript framework

1. Verify everything is working. Execute the following command to build, start the local web server and test the web part in the local workbench:

    ```shell
    gulp serve
    ```

1. When the browser loads the local workbench, select the **Add a new web part** control...

    ![Screenshot of the SharePoint workbench](../../Images/EditPropPane-TestWP-01.png)

   ...and select the **HelloPropertyPane** web part to add the web part to the page:

    ![Screenshot of HelloPropertyPane web part in the SharePoint workbench](../../Images/EditPropPane-TestWP-02.png)

1. Select the **edit web part** control to the side of the web part to display the property pane:

    ![Screenshot of HelloPropertyPane's property pane](../../Images/EditPropPane-TestWP-03.png)

## Add New Properties to the web part

With a working web part, the next step is to customize the property pane experience.

1. Create two new properties that will be used in the web part and property pane:
    1. Open the file **src\webparts\helloPropertyPane\HelloPropertyPaneWebPart.ts**
    1. Locate the interface `IHelloPropertyPaneWebPartProps` after the `import` statements. Add the following two properties to the interface:

    ```ts
    myContinent: string;
    numContinentsVisited: number;
    ```

1. Update the web part rendering to display the values of these two properties:
    1. within the `HelloPropertyPaneWebPart` class, locate the `render()` method.
    1. Within the `render()` method, locate the following line in the HTML output:

        ```html
        <p class="${ styles.description }">${escape(this.properties.description)}</p>
        ```

    1. Add the following two lines after the line you just located:

        ```html
        <p class="${ styles.description }">Continent where I reside: ${escape(this.properties.myContinent)}</p>
        <p class="${ styles.description }">Number of continents I've visited: ${this.properties.numContinentsVisited}</p>
        ```

    At the moment the web part will render a blank string and undefined for these two fields as nothing is present in their values:

      ![Screenshot of HelloPropertyPane with no values](../../Images/EditPropPane-AddProps-01.png)

    This can be addressed by setting the default values of properties when a web part is added to the page.

1. Set the default property values:
    1. Open the file **src\webparts\helloPropertyPane\HelloPropertyPaneWebPart.manifest.json**
    1. Locate the following section in the file: `preconfiguredEntries[0].properties.description`
    1. Add a comma after the `description` property's value.
    1. Add the following two lines after the `description` property:

        ```json
        "myContinent": "North America",
        "numContinentsVisited": 4
        ```

1. Updates to the web part's manifest file will not be picked up until you restart the local web server.
    1. In the command prompt, press <kbd>CTRL+C</kbd> to stop the local web server.
    1. Rebuild and restart the local web server by executing the command `gulp serve`.
    1. When the SharePoint workbench loads, add the web part back tot he page to see the properties.

        ![Screenshot of HelloPropertyPane with no values](../../Images/EditPropPane-AddProps-02.png)

## Extend the Property Pane

Now that the web part has two new custom properties, the next step is to extend the property pane to allow users to edit the values.

1. Add a new text control to the property pane, connected to the **myContinent** property:
    1. Open the file **src\webparts\helloPropertyPane\HelloPropertyPaneWebPart.ts**
    1. Locate the method `getPropertyPaneConfiguration` & within it, locate the `groupFields` array.
    1. Add a comma after the existing `PropertyPaneTextField()` call.
    1. Add the following code after the comma:

        ```ts
        PropertyPaneTextField('myContinent', {
          label: 'Continent where I currently reside'
        })
        ```

1. If the local web server is not running, start it by executing `gulp serve`.

    Once the SharePoint workbench is running again, add the web part to the page and open the property pane.

    Notice you can see the property and text box in the property pane. Any edits to the field will automatically update the web part:

    ![Screenshot of HelloPropertyPane with a new field](../../Images/EditPropPane-AddProps-03.png)

1. Now add a slider control to the property pane, connected to the **numContinentsVisited** property:
    1. In the **HelloPropertyPaneWebPart.ts**, at the top of the file, add a `PropertyPaneSlider` reference to the existing `import` statement for the `@microsoft/sp-webpart-base` package.
    1. Scroll down to the method `getPropertyPaneConfiguration` & within it, locate the `groupFields` array.
    1. Add a comma after the last `PropertyPaneTextField()` call, add the following code:

        ```ts
        PropertyPaneSlider('numContinentsVisited', {
          label: 'Number of continents I\'ve visited',
          min: 1, max: 7, showValue: true,
        })
        ```

    1. Notice the property pane now has a slider control to control the number of continents you have visited:

        ![Screenshot of HelloPropertyPane with a new field](../../Images/EditPropPane-AddProps-04.png)

## Add Control Validation

In a previous step the user was given a property where they could enter the name of the continent in which they live. Add validation logic to ensure they enter a valid continent name.

1. In the **HelloPropertyPaneWebPart.ts**, add the following method to the `HelloPropertyPaneWebPart` class. This method takes a string as an input and returns a string. This allows you to do custom validation logic. If this method returns an empty string, the value is considered valid; otherwise, the returned string is used as the error message.

    ```ts
    private validateContinents(textboxValue: string): string {
      const validContinentOptions: string[] = ['africa', 'antarctica', 'asia', 'australia', 'europe', 'north america', 'south america'];
      const inputToValidate: string = textboxValue.toLowerCase();

      return (validContinentOptions.indexOf(inputToValidate) === -1)
        ? 'Invalid continent entry; valid options are "Africa", "Antarctica", "Asia", "Australia", "Europe", "North America", and "South America"'
        : '';
    }
    ```

1. Wire the validation method to the text field control previously added.
    1. Locate the text field that is associated with the **myContinent** property.
    1. Add the following code to the options object passed into the `PropertyPaneTextField()` call as the second parameter: `onGetErrorMessage: this.validateContinents.bind(this)`.

        The text field control should now look like the following code: 

        ```ts
        PropertyPaneTextField('myContinent', {
          label: 'Continent where I currently reside',
          onGetErrorMessage: this.validateContinents.bind(this)
        }),
        ```

1. If the local web server is not running, start it by executing `gulp serve`.

    Once the SharePoint workbench is running again, add the web part to the page and open the property pane.

1. Enter the name of a non-existent continent to test the validation logic:

    ![Screenshot of HelloPropertyPane with validation logic applied](../../Images/EditPropPane-AddProps-05.png)

    Notice the property value is not updated when the control's input is invalid.
