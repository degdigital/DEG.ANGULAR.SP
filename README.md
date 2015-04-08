# AngularSP
AngularSP is an Angular module containing a factory and associated directives for interacting with SharePoint.

## Usage

### Building your List Definition
The module relies on a list definition JSON object to be created for the "spListFactory" to obtain the list information for use by the factory and directives.  This JSON object should follow these rules:
* Each list should have an object in the root object keyed with the list's internal name.
  * Each list object should have a "Display Name" if the list's display name is different from it's internal name.
  * Each list object *must* have a "Columns" object that contains the columns that will be present on the form.
    * The "Columns" object should contain one or more column objects keyed with the column's internal name.
    * Each column object *must* have a "Type" and can have an "Input Type" to override the column type default.  Available types and their available input types are:
      * "text"
      * "choice"
      * "multiText"
      * "link"
      * "number"
      * "currency"
      * "multiChoice"
      * "lookup"
      * "multiLookup"
      * "yesNo"
      * "person"
      * "multiPerson"
      * "metadata"
      * "multiMetadata"

### Loading Module
The "angularSP" module needs to be loaded on your app.

#### Example
```
angular.module("app", ["angularSP"]);
```

### Injecting Factory
The "spListFactory" needs to be injected into your controller.

#### Example
```
angular.module("app").controller("ctrl", ["$scope", "spListFactory", function ($scope, spListFactory) { }]);
```

### Initializing Factory