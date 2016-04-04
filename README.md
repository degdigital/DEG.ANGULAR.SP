# AngularSP
**Version 1.1**

AngularSP is an Angular module containing a provider and associated directives for interacting with SharePoint.

## Usage

### Setting Configuration with the List Definition
The module relies on a list definition JSON object to be set in the config stage using the "defineLists" function.  This JSON object should follow these rules:
* Each list must have an object in the root object keyed with the list's internal name.
  * Each list object should have a "DisplayName" if the list's display name is different from it's internal name.
  * Each list object should have a "ContextUrl" if the list's context is different then the current context where the app is executing.

#### Example
```
app.config(['spFormProviderProvider', function (spFormProvider) {
    var listDefinitions = {
        'AngularSPTestList': {
            'DisplayName': 'AngularSP Test List'
        },
        'AngularSPGridList': {
            'DisplayName': 'AngularSP Grid List'
        }
    };
    spFormProvider.defineLists(listDefinitions);
}]);
```

### Loading Module
The "angularSP" module needs to be loaded on your app.

#### Example
```
angular.module("app", ["angularSP"]);
```

### Injecting Factory
The "spListProvider" needs to be injected into your controller.

#### Example
```
angular.module("app").controller("ctrl", ["$scope", "spFormProvider", function ($scope, spFormProvider) { }]);
```

### Initializing Provider
The "spFormProvider" gets initialized when injected into the controller.  The injected object contains a promise called "listInformationPromise" that returns the list information once it is retrieved and initialization is complete.  Generally, you will want wait for the "listInformationPromise" promise to return before executing any controller code.

#### Example
```
app.controller('testFormController', ['$scope', '$log', 'spFormProvider', function($scope, $log, spFormProvider) {
	spFormProvider.listInformationPromise.then(function (listsInfo) {
		$scope.listsInfo = listsInfo;
		...
	}, function (reason) {
		$log.error(reason);
	});
}]);
```

### Provider Calls
Once the provider is initialized, you can use the following methods to interact with your SharePoint lists:
* createListItem(listName, itemProperties)
  * listName is the internal list name which *must* match up with the key for the list in the list definition JSON.
  * itemProperties is an object keyed off from column internal names that *must* match up to those in the list definition JSON.
  * returns a promise that resolves the numerical Id of the newly created item.
* createListItems(listName, itemsProperties)
  * listName is the internal list name which *must* match up with the key for the list in the list definition JSON.
  * itemsProperties is an array of objects, each keyed off from column internal names that *must* match up to those in the list definition JSON.
  * return a promise that resolves an array of numerical Ids for the newly created items.
* getListItem(listName, itemId)
  * listName is the internal list name which *must* match up with the key for the list in the list definition JSON.
  * itemId is the numerical id of the list item that you are requesting.
  * returns a promise that resolves an object keyed with the column names which map to their values.
* getListItems(listName, parameters)
  * listName is the internal list name which *must* match up with the key for the list in the list definition JSON.
  * parameters is an object that *currently* only supports a LookupColumn, which is the internal column name for a lookup type column in the list, and LookupValue, which is simply the value the query needs to match for the list items.
  * returns a promise that resolves an array of objects, each keyed with the column names which map to their values.
* updateListItem(listName, itemProperties, overrideConflict)
  * listName is the internal list name which *must* match up with the key for the list in the list definition JSON.
  * itemProperties is an object keyed off from column internal names that *must* match up to those in the list definition JSON.
  * overrideConflict is a boolean value that specifies if you want the update to override an update conflict (usually caused by another user or workflow saving the form while this one is open).
  * returns a promise that resolves when the item is successfully updated.
* deleteListItem(listName, itemId)
  * listName is the internal list name which *must* match up with the key for the list in the list definition JSON.
  * itemId is the numerical id of the list item that you are deleting.
  * returns a promise that resolves when the item is successfully deleted.
* commitListItems(listName, parameters, itemsProperties)
  * listName is the internal list name which *must* match up with the key for the list in the list definition JSON.
  * parameters is an object that *currently* only supports a LookupColumn, which is the internal column name for a lookup type column in the list, and LookupValue, which is simply the value the query needs to match for the list items.
  * itemsProperties is an array of objects, each keyed off from column internal names that *must* match up to those in the list definition JSON.
  * returns a promise that resolves when all item changes have been commited.
* getServerRelativeUrl()
  * returns the server relative url of the current site.
  
### Directives
On the view, there are two attribute directives that can be used to specify where you would like the input to appear.
* ngsp-column
  * specifies the column info obtained during the initialization.
  * additionally requires an ng-model to be defined.
* ngsp-list
  * specifies the column info obtained during the initialization.
  * additionally requires an ng-model to be defined.
  * additionally requires an ngsp-template-url to be defined for the create/update modal.
  
#### Example
```
<div data-ngsp-column="listsInfo.AngularSPTestList.Columns.Title" data-ng-model="AngularSPTestList.Title"></div>
<script type="text/ng-template" id="AngularSPGridListTemplate">
    <div class="modal-body">
        <div data-ngsp-column="columns.Title" data-ng-model="item.Title"></div>
    </div>
    <div class="modal-footer">
        <input type="button" ng-click="submit(item)" value="Submit" />
        <input type="button" ng-click="cancel()" value="Cancel" />
    </div>
</script>
<div data-ngsp-list="listsInfo.AngularSPGridList" data-ng-model="AngularSPGridList" data-ngsp-template-url="AngularSPGridListTemplate"></div>
```