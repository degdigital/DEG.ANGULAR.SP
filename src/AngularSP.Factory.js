﻿
var app = app || angular.module("angularSP", ['ui.bootstrap']);

(function (ng, $) {

    app.factory("spListFactory", ['$rootScope', '$q', spListFactory]);

    function spListFactory($rootScope, $q) {

        var service = {
            initFactory: initFactory,
            createListItem: createListItem,
            createListItems: createListItems,
            getListItem: getListItem,
            getListItems: getListItems,
            updateListItem: updateListItem,
            deleteListItem: deleteListItem,
            getServerRelativeUrl: getServerRelativeUrl
        }

        var serverRelativeUrl;

        var columnDefinitions;
        var contentTypeDefinitions;
        var listDefinitions;

        var columnsInfo;
        var cTypesInfo;
        var listsInfo;

        return service;

        function getServerRelativeUrl() {
            return serverRelativeUrl;
        }

        function initFactory(columnDefs, cTypeDefs, listDefs) {

            columnDefinitions = columnDefs;
            columnDefinitions = cTypeDefs;
            listDefinitions = listDefs;

            columnsInfo = ng.copy(columnDefs);
            cTypesInfo = ng.copy(cTypeDefs);
            listsInfo = ng.copy(listDefs);

            cascadeColumnProperties(columnsInfo, cTypesInfo, listsInfo);

            var initDeferred = $q.defer();

            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                SP.SOD.executeFunc('sp.taxonomy.js', 'SP.Taxonomy.TaxonomySession', function () {
                    SP.SOD.executeFunc('SP.UserProfiles.js', 'SP.UserProfiles.PeopleManager', sharePointReady);
                });
            });

            return initDeferred.promise;

            function sharePointReady() {

                var context = SP.ClientContext.get_current();
                var web = context.get_web();

                context.load(web, "ServerRelativeUrl");

                //var columnFieldData = loadSiteColumnProperties(columnDefs);
                var listFieldData = loadListInfoProperties(listDefs);

                context.executeQueryAsync(success, failure);

                function success() {
                    serverRelativeUrl = web.get_serverRelativeUrl();
                    //setSiteColumnProperties(columnsInfo, columnFieldData);
                    var fieldLookups = setListInfoProperties(listsInfo, listFieldData);

                    context.executeQueryAsync(lookupsSuccess, lookupsFailure);

                    function lookupsSuccess() {
                        fieldLookups.forEach(function (fieldLookup) {
                            fieldLookup.Callback(fieldLookup.Column, fieldLookup.RawResults);
                        });
                        $rootScope.listsInfo = listsInfo;
                        initDeferred.resolve(listsInfo);
                    }

                    function lookupsFailure(sender, args) {
                        console.log(args.get_message());
                        initDeferred.reject(args.get_message());
                    }

                }

                function failure(sender, args) {
                    console.log(args.get_message());
                    initDeferred.reject(args.get_message());
                }
            }
        }

        function createListItem(listName, itemProps) {
            var createDeferred = $q.defer();
            console.log(listName);
            var listInfo;
            if (listsInfo !== undefined) {
                if (listsInfo[listName] !== undefined) {
                    listInfo = listsInfo[listName];
                    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                        SP.SOD.executeFunc('sp.taxonomy.js', 'SP.Taxonomy.TaxonomySession', function () {
                            SP.SOD.executeFunc('SP.UserProfiles.js', 'SP.UserProfiles.PeopleManager', sharePointReady);
                        });
                    });
                }
            }
            else {

            }

            return createDeferred.promise;

            function sharePointReady() {
                //console.log(listsInfo);
                //console.log(listName);
                //var listInfo = listsInfo[listName];
                //console.log(listInfo);
                var listName = listInfo.DisplayName;

                var context = SP.ClientContext.get_current();
                var web = context.get_web();
                var spList = web.get_lists().getByTitle(listName);
                var itemCreationInfo = new SP.ListItemCreationInformation();
                var listItem = spList.addItem(itemCreationInfo);
                
                var columns = listInfo.Columns;
                for (var columnName in columns) {
                    if (columns.hasOwnProperty(columnName)) {
                        var columnInfo = columns[columnName];
                        if (itemProps[columnName] !== undefined) {
                            var inputValue = itemProps[columnName];
                            var itemValue = getListItemValue(columnInfo, inputValue);
                            //console.log(columnName);
                            //console.log(itemValue);
                            listItem.set_item(columnName, itemValue);
                        }
                    }
                }
                listItem.update();
                context.load(listItem);
                context.executeQueryAsync(function () {
                    createDeferred.resolve(listItem.get_id());
                }, function (sender, args) {
                    createDeferred.reject(args.get_message());
                });
            }
        }

        function createListItems(listName, itemsProps) {
            var createDeferred = $q.defer();



            return createDeferred.promise;
        }

        function getListItem(listName, itemId) {
            var listItemDeferred = $q.defer();

            var listInfo;
            var listItemId;

            if (listsInfo !== undefined) {
                if (listsInfo[listName] !== undefined) {
                    listInfo = listsInfo[listName];
                    listItemId = itemId;
                    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                        SP.SOD.executeFunc('sp.taxonomy.js', 'SP.Taxonomy.TaxonomySession', function () {
                            SP.SOD.executeFunc('SP.UserProfiles.js', 'SP.UserProfiles.PeopleManager', sharePointReady);
                        });
                    });
                }
            }
            else {

            }

            return listItemDeferred.promise;

            function sharePointReady() {
                var listName = listInfo.DisplayName;

                var context = SP.ClientContext.get_current();
                var web = context.get_web();
                var spList = web.get_lists().getByTitle(listName);
                var spItem = spList.getItemById(listItemId);
                context.load(spItem);
                context.executeQueryAsync(function () {
                    var columns = listInfo.Columns;
                    var spItemValues = spItem.get_fieldValues();
                    var inputValues = {};
                    var peopleLookups = [];
                    for (var columnName in columns) {
                        if (columns.hasOwnProperty(columnName)) {
                            var columnInfo = columns[columnName];
                            if (spItemValues[columnName] !== undefined) {
                                var spItemValue = spItemValues[columnName];
                                if (spItemValue !== null) {
                                    var inputValue = getInputValue(columnInfo, spItemValue, peopleLookups);
                                    inputValues[columnName] = inputValue;
                                }
                                //console.log(columnName);
                                //console.log(inputValues[columnName]);
                            }
                        }
                    }
                    if (peopleLookups.length > 0) {
                        context.executeQueryAsync(function () {
                            peopleLookups.forEach(function (peopleLookup) {
                                peopleLookup.Callback(peopleLookup.Values, peopleLookup.RawResults);
                            });
                            listItemDeferred.resolve(inputValues);
                        }, function (sender, args) {
                            listItemDeferred.reject(args.get_message());
                        });
                    }
                    else {
                        listItemDeferred.resolve(inputValues);
                    }
                }, function (sender, args) {
                    listItemDeferred.reject(args.get_message());
                });
            }
        }

        function getListItems(listName, params) {

        }

        function updateListItem(listName, itemProps, overrideConflict) {

        }

        function deleteListItem(listName, itemId) {
            var deleteDeferred = $q.defer();

            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            var list = web.get_lists().getByTitle(libraryName);
            var item = list.getItemById(itemId);
            item.deleteObject();
            context.executeQueryAsync(success, failure);

            return deleteDeferred.promise;

            function success() {
                deleteDeferred.resolve();
            }

            function failure(sender, args) {
                deleteDeferred.reject(sender, args);
            }
        }

        /**/

        function getListItemValue(columnInfo, inputValue) {
            var itemValue = null;
            switch (columnInfo.Type) {
                case "text":
                case "choice":
                case "multiText":
                case "link":
                case "number":
                case "currency":
                    itemValue = inputValue;
                    break;
                case "multiChoice":
                    itemValue = [];
                    for (var value in inputValue) {
                        if (inputValue.hasOwnProperty(value)) {
                            if (inputValue[value]) {
                                itemValue.push(value);
                            }
                        }
                    }
                    break;
                case "lookup":
                    itemValue = new SP.FieldLookupValue();
                    itemValue.set_lookupId(inputValue);
                    break;
                case "multiLookup":
                    itemValue = [];
                    for (var value in inputValue) {
                        if (inputValue.hasOwnProperty(value)) {
                            if (inputValue[value]) {
                                var lookupValue = new SP.FieldLookupValue();
                                lookupValue.set_lookupId(value);
                                itemValue.push(lookupValue);
                            }
                        }
                    }
                    break;
                case "yesNo":
                    if (inputValue) {
                        itemValue = 1;
                    }
                    else {
                        itemValue = 0;
                    }
                    break;
                case "person":
                    if (inputValue.length === 1) {
                        var person = inputValue[0];
                        itemValue = SP.FieldUserValue.fromUser(person.Key);
                    }
                    else if (inputValue.length !== 0) {
                        //should be multiPerson
                    }
                    break;
                case "multiPerson":
                    itemValue = [];
                    for (var i = 0; i < inputValue.length; i++) {
                        var person = inputValue[i];
                        itemValue.push(SP.FieldUserValue.fromUser(person.Key));
                    }
                    break;
                case "metadata":
                    var metadataItem = columnInfo.TermsById[inputValue];
                    itemValue = metadataItem.label + "|" + metadataItem.id;
                    break;
                case "multiMetadata":
                    var metadataValues = [];
                    for (var value in inputValue) {
                        if (inputValue.hasOwnProperty(value)) {
                            if (inputValue[value]) {
                                var metadataItem = columnInfo.TermsById[value];
                                metadataValues.push("-1");
                                metadataValues.push(metadataItem.label + "|" + metadataItem.id);
                            }
                        }
                    }
                    itemValue = metadataValues.join(";#");
                    break;
                case "dateTime":
                    itemValue = inputValue.toISOString();
                    break;
                default:
                    break;
            }
            return itemValue;
        }

        function getInputValue(columnInfo, itemValue, peopleLookups) {
            var inputValue;
            switch (columnInfo.Type) {
                case "text":
                case "choice":
                case "multiText":
                case "link":
                case "number":
                case "currency":
                    inputValue = itemValue;
                    break;
                case "multiChoice":
                    inputValue = {};
                    for (var i = 0; i < itemValue.length; i++) {
                        var value = itemValue[i];
                        inputValue[value] = true;
                    }
                    break;
                case "lookup":
                    inputValue = itemValue.get_lookupId();
                    break;
                case "multiLookup":
                    inputValue = {};
                    for (var i = 0; i < itemValue.length; i++) {
                        var value = itemValue[i].get_lookupId();
                        inputValue[value] = true;
                    }
                    break;
                case "yesNo":
                    if (itemValue === 1) {
                        inputValue = true;
                    }
                    else {
                        inputValue = false;
                    }
                    break;
                case "person":
                    inputValue = [];
                    var context = SP.ClientContext.get_current();
                    var web = context.get_web();
                    var rawValue = web.getUserById(itemValue.get_lookupId());
                    context.load(rawValue);
                    peopleLookups.push({ Values: inputValue, RawResults: rawValue, Callback: setPersonInputValue });
                    break;
                case "multiPerson":
                    inputValue = [];
                    var context = SP.ClientContext.get_current();
                    var web = context.get_web();
                    var rawValues = [];
                    for (var i = 0; i < itemValue.length; i++) {
                        var rawValue = web.getUserById(itemValue[i].get_lookupId());
                        context.load(rawValue);
                        rawValues.push(rawValue);
                    }
                    peopleLookups.push({ Values: inputValue, RawResults: rawValues, Callback: setPersonInputValue });
                    break;
                case "metadata":
                    inputValue = itemValue.get_termGuid()
                    break;
                case "multiMetadata":
                    inputValue = {};
                    for (var i = 0; i < itemValue.get_count() ; i++) {
                        var value = itemValue.getItemAtIndex(i);
                        inputValue[value.get_termGuid()] = true;
                    }
                    break;
                case "dateTime":
                    inputValue = new Date(itemValue);
                    break;
                default:
                    break;
            }
            return inputValue;
        }

        function setPersonInputValue(values, rawResults) {
            if (rawResults.length > 0) {
                rawResults.forEach(addPersonToValues);
            }
            else if(rawResults !== null) {
                addPersonToValues(rawResults);
            }
            function addPersonToValues(result) {
                var person = {
                    AutoFillDisplayText: result.get_title(),
                    AutoFillKey: result.get_loginName(),
                    Description: result.get_email(),
                    DisplayText: result.get_title(),
                    EntityType: "User",
                    IsResolved: true,
                    Key: result.get_loginName(),
                    Resolved: true
                }
                values.push(person);
            }
        }

        function cascadeColumnProperties(columnsInfo, cTypesInfo, listsInfo) {
            cascadeListsInfoProperties(columnsInfo, listsInfo);
        }

        function cascadeListsInfoProperties(columnsInfo, listsInfo) {
            for (var listName in listsInfo) {
                if (listsInfo.hasOwnProperty(listName)) {
                    var listColumns = listsInfo[listName].Columns;
                    if (listColumns !== undefined) {
                        cascadeListInfoColumnProperties(columnsInfo, listColumns);
                    }
                }
            }
        }

        function cascadeListInfoColumnProperties(columnsInfo, listColumns) {
            for (var columnName in listColumns) {
                if (listColumns.hasOwnProperty(columnName)) {
                    var column = listColumns[columnName];
                    var columnInfo = columnsInfo[columnName];
                    if (columnInfo !== undefined) {
                        cascadeProperties(columnInfo, column);
                    }
                }
            }
        }

        function cascadeProperties(columnFrom, columnTo) {
            for (var propertyName in columnFrom) {
                if (columnFrom.hasOwnProperty(propertyName)) {
                    if (columnTo[propertyName] === undefined) {
                        columnTo[propertyName] = columnFrom[propertyName];
                    }
                }
            }
        }

        function loadSiteColumnProperties(columnDefs) {
            var fieldData = [];
            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            var fields = web.get_fields();
            for (var columnName in columnDefs) {
                if (columnsInfo.hasOwnProperty(columnName)) {
                    var field = fields.getByInternalNameOrTitle(columnName);
                    context.load(field);
                    fieldData.push(field);
                }
            }
            return fieldData;
        }

        function setSiteColumnProperties(columnsInfo, fieldData) {
            var context = SP.ClientContext.get_current();
            for (var i = 0; i < fieldData.length; i++) {
                var field = fieldData[i];
                var internalName = field.get_internalName();
                if (columnsInfo[internalName] !== undefined) {
                    var column = columnsInfo[internalName];
                    setColumnInfo(column, field, 0);
                }
                else {
                    //Very strange error.
                    console.log("Field data mismatch.");
                }
            }
        }

        function setColumnInfo(column, field, fieldLookups, lookupCount) {
            var context = SP.ClientContext.get_current();
            column.Title = field.get_title();
            column.Id = field.get_id().toString();
            column.IsRequired = field.get_required();
            column.SchemaXml = field.get_schemaXml();
            column.DefaultValue = field.get_defaultValue();
            //column.Value = field.get_defaultValue();
            column.Scope = field.get_scope();
            column.FieldType = field.get_fieldTypeKind();
            switch (column.Type) {
                case "text":
                    if (column.FieldType === 2) {
                        if (column.InputType === undefined) {
                            column.InputType = "text";
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    break;
                case "multiText":
                    if (column.FieldType === 3) {
                        var multiLineTextField = context.castTo(field, SP.FieldMultiLineText);
                        column.IsRichText = multiLineTextField.get_richText();
                        if (column.InputType === undefined) {
                            if (column.IsRichText) {
                                column.InputType = "nicEdit";
                            }
                            else {
                                column.InputType = "textBox";
                            }
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    break;
                case "choice":
                    if (column.FieldType === 6) {
                        if (column.InputType === undefined) {
                            column.InputType = "dropDown";
                        }
                        var choiceField = context.castTo(field, SP.FieldChoice);
                        column.Choices = choiceField.get_choices();
                        column.FillInChoice = choiceField.get_fillInChoice();
                        column.EditFormat = choiceField.get_editFormat();
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    break;
                case "multiChoice":
                    if (column.FieldType === 15) {
                        if (column.InputType === undefined) {
                            column.InputType = "checkbox";
                        }
                        var multiChoiceField = context.castTo(field, SP.FieldMultiChoice);
                        column.Choices = multiChoiceField.get_choices();
                        column.FillInChoice = multiChoiceField.get_fillInChoice();
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    break;
                case "lookup":
                    if (column.FieldType === 7) {
                        if (column.InputType === undefined) {
                            column.InputType = "dropDown";
                        }
                        var lookupField = context.castTo(field, SP.FieldLookup);
                        var multiLookup = lookupField.get_allowMultipleValues();
                        if (!multiLookup) {
                            column.LookupField = lookupField.get_lookupField();
                            column.LookupList = lookupField.get_lookupList();
                            //column.LookupWeb = lookupField.get_lookupWebId().toString();
                            var rawlookupItems = loadLookupOptions(column);
                            fieldLookups.push({ Column: column, RawResults: rawlookupItems, Callback: setLookupOptions });
                            lookupCount++;
                        }
                        else {
                            console.log(field.get_title() + ": column type mismatch.");
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    break;
                case "multiLookup":
                    if (column.FieldType === 7) {
                        if (column.InputType === undefined) {
                            column.InputType = "checkbox";
                        }
                        var lookupField = context.castTo(field, SP.FieldLookup);
                        var multiLookup = lookupField.get_allowMultipleValues();
                        if (multiLookup) {
                            column.LookupField = lookupField.get_lookupField();
                            column.LookupList = lookupField.get_lookupList();
                            //column.LookupWeb = lookupField.get_lookupWebId().toString();
                            var rawlookupItems = loadLookupOptions(column);
                            fieldLookups.push({ Column: column, RawResults: rawlookupItems, Callback: setLookupOptions });
                            lookupCount++;
                        }
                        else {
                            console.log(field.get_title() + ": column type mismatch.");
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    break;
                case "yesNo":
                    if (column.FieldType === 8) {
                        if (column.InputType === undefined) {
                            column.InputType = "checkbox";
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    break;
                case "person":
                    if (column.FieldType === 20) {
                        if (column.InputType === undefined) {
                            column.InputType = "peoplePicker";
                        }
                        var userField = context.castTo(field, SP.FieldUser);
                        column.SelectionGroup = userField.get_selectionGroup();
                        column.PeopleOnly = (userField.get_selectionMode() === 0);
                        var multiUsers = userField.get_allowMultipleValues();
                        if (!multiUsers) {
                            lookupCount++;
                        }
                        else {
                            console.log(field.get_title() + ": column type mismatch.");
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //peopleOnly/peopleAndGroups
                    //chooseFrom(group)
                    //showField
                    break;
                case "multiPerson":
                    if (column.FieldType === 20) {
                        if (column.InputType === undefined) {
                            column.InputType = "peoplePicker";
                        }
                        var userField = context.castTo(field, SP.FieldUser);
                        column.SelectionGroup = userField.get_selectionGroup();
                        column.PeopleOnly = (userField.get_selectionMode() === 0);
                        var multiUsers = userField.get_allowMultipleValues();
                        if (multiUsers) {
                            lookupCount++;
                        }
                        else {
                            console.log(field.get_title() + ": column type mismatch.");
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //peopleOnly/peopleAndGroups
                    //chooseFrom(group)
                    //showField
                    break;
                case "metadata":
                    if (column.FieldType === 0) {
                        if (column.InputType === undefined) {
                            column.InputType = "dropDown";
                        }
                        var schemaFieldType = parseMetadataSchema(column);
                        if (schemaFieldType === "TaxonomyFieldType") {
                            var rawTermSet = loadMetadataTerms(column);
                            fieldLookups.push({ Column: column, RawResults: rawTermSet, Callback: setMetadataTerms });
                            lookupCount++;
                        }
                        else {
                            console.log(field.get_title() + ": column type mismatch.");
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //displayValue
                    //allowFillIn
                    //default
                    break;
                case "multiMetadata":
                    if (column.FieldType === 0) {
                        if (column.InputType === undefined) {
                            column.InputType = "checkbox";
                        }
                        var schemaFieldType = parseMetadataSchema(column);
                        if (schemaFieldType === "TaxonomyFieldTypeMulti") {
                            var rawTermSet = loadMetadataTerms(column);
                            fieldLookups.push({ Column: column, RawResults: rawTermSet, Callback: setMetadataTerms });
                            lookupCount++;
                        }
                        else {
                            console.log(field.get_title() + ": column type mismatch.");
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //displayValue
                    //allowFillIn
                    //default
                    break;
                case "link":
                    if (column.FieldType === 11) {
                        if (column.InputType === undefined) {
                            column.InputType = "text";
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //format (hyperlink/picture)
                    break;
                case "number":
                    if (column.FieldType === 9) {
                        if (column.InputType === undefined) {
                            column.InputType = "text";
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //minmax
                    //decimalplaces
                    //showaspercentage
                    break;
                case "currency":
                    if (column.FieldType === 10) {
                        if (column.InputType === undefined) {
                            column.InputType = "text";
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //minmax
                    //decimalplaces
                    //format
                    break;
                case "dateTime":
                    if (column.FieldType === 4) {
                        if (column.InputType === undefined) {
                            column.InputType = "dateTime";
                        }
                    }
                    else {
                        console.log(field.get_title() + ": column type mismatch.");
                    }
                    //format (dateOnly/dateTime) (standard/friendly)
                    break;
                default:
                    //throw error.
                    break;
            }
            return lookupCount;
        }

        function loadListInfoProperties(listDefs) {
            var fieldData = [];
            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            for (var listName in listDefs) {
                if (listDefs.hasOwnProperty(listName)) {
                    var listDef = listDefs[listName];
                    var listTitle = listName;
                    if (listDef.DisplayName !== undefined) {
                        listTitle = listDef.DisplayName;    
                    }
                    var list = web.get_lists().getByTitle(listTitle);
                    if (listDef.Columns !== undefined) {
                        loadListFields(context, list, listDef.Columns, fieldData);
                    }
                    else {
                        //Report missing column defenitions Error.
                    }
                }
            }
            return fieldData;
        }

        function loadListFields(context, list, listColumns, fieldData) {
            var fields = list.get_fields();
            for (var columnName in listColumns) {
                if (listColumns.hasOwnProperty(columnName)) {
                    var field = fields.getByInternalNameOrTitle(columnName);
                    context.load(field);
                    fieldData.push(field);
                }
            }
        }

        function loadLookupOptions(column) {
            var query = new SP.CamlQuery();
            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            var list = web.get_lists().getById(column.LookupList);
            var lookupItems = list.getItems(query);
            context.load(lookupItems);
            return lookupItems;
        }

        function loadMetadataTerms(column) {
            var context = SP.ClientContext.get_current();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
            var termStores = taxSession.get_termStores();
            var termStore = termStores.getById(column.SspId);
            var termSet = termStore.getTermSet(column.TermSetId);
            var terms = termSet.getAllTerms();
            context.load(terms);
            return terms;
        }

        function setLookupOptions(column, lookupItems) {
            column.LookupItems = [];
            column.LookupItemsById = {};
            column.LookupItemsByLabel = {};
            var lookupItemsEnum = lookupItems.getEnumerator();
            while (lookupItemsEnum.moveNext()) {
                var lookupItem = lookupItemsEnum.get_current();
                var lookupItemValues = lookupItem.get_fieldValues();
                var lookupItemId = lookupItemValues.ID;
                var lookupItemLabel = lookupItemValues[column.LookupField];
                var lookupItemObj = { id: lookupItemId, label: lookupItemLabel };
                column.LookupItems.push(lookupItemObj);
                column.LookupItemsById[lookupItemId] = lookupItemObj;
                column.LookupItemsByLabel[lookupItemLabel] = lookupItemObj;
            }
        }

        function setMetadataTerms(column, terms) {
            column.Terms = [];
            column.TermsById = {};
            column.TermsByLabel = {};
            var termEnum = terms.getEnumerator();
            while (termEnum.moveNext()) {
                var currentTerm = termEnum.get_current();
                var currentTermId = currentTerm.get_id().toString();
                var currentTermLabel = currentTerm.get_name();
                var termObj = { id: currentTermId, label: currentTermLabel }
                column.Terms.push(termObj);
                column.TermsById[currentTermId] = termObj;
                column.TermsByLabel[currentTermLabel] = termObj;
            }
        }

        function setListInfoProperties(listsInfo, fieldData) {
            var fieldLookups = [];
            for (var i = 0; i < fieldData.length; i++) {
                var field = fieldData[i];
                var internalName = field.get_internalName();
                var scope = field.get_scope();
                var listTitle = scope.split("/").pop();
                var listInfo = listsInfo[listTitle];
                if (listInfo !== undefined && listInfo.Columns !== undefined) {
                    var column = listInfo.Columns[internalName];
                    if (listInfo.LookupCount === undefined) {
                        listInfo.LookupCount = 0;
                    }
                    listInfo.LookupCount = setColumnInfo(column, field, fieldLookups, listInfo.LookupCount);
                }
                else {
                    //May be a site column instead of a list column.  Although this shouldn't happen.
                }
            }
            return fieldLookups;
        }

        //function setListInfoProperties(listsInfo) {
        //    var fieldPromises = [];
        //    for (var listName in listsInfo) {
        //        if (listsInfo.hasOwnProperty(listName)) {
        //            var list = listsInfo[listName];
        //            var listTitle = listName;
        //            if (list.DisplayName !== undefined) {
        //                listTitle = list.DisplayName;
        //            }
        //            if (list.Columns !== undefined) {
        //                for (var columnName in list.Columns) {
        //                    if (list.Columns.hasOwnProperty(columnName)) {
        //                        var column = list.Columns[columnName];
        //                        if (column.Type !== undefined && column.InputType !== undefined) {
        //                            var fieldPromise = getListField(listTitle, columnName);
        //                            fieldPromise.then(function (field) {

        //                            }, function (reason) {

        //                            });
        //                            switch (column.Type) {
        //                                case "text":
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(field);
        //                                        column.DefaultValue = field.get_defaultValue();
        //                                        column.IsRequired = field.get_required();
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    break;
        //                                case "multiText":
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(field);
        //                                        column.IsRequired = field.get_required();
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    break;
        //                                case "choice":
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //choices
        //                                    //default
        //                                    //displayAs
        //                                    break;
        //                                case "multiChoice":
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //choices
        //                                    //default
        //                                    break;
        //                                case "lookup":
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //choices
        //                                    break;
        //                                case "multiLookup":
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //choices
        //                                    break;
        //                                case "yesNo":
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                        column.DefaultValue = field.get_defaultValue();
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    break;
        //                                case "person":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //peopleOnly/peopleAndGroups
        //                                    //chooseFrom(group)
        //                                    //showField
        //                                    break;
        //                                case "multiPerson":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //peopleOnly/peopleAndGroups
        //                                    //chooseFrom(group)
        //                                    //showField
        //                                    break;
        //                                case "metadata":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //choices
        //                                    //displayValue
        //                                    //allowFillIn
        //                                    //default
        //                                    var termsPromise = getTermsForField(listName, columnName);
        //                                    termsPromise.then(function (terms) {
        //                                        setTermsForField(column, terms);
        //                                    }, function (reason) {

        //                                    });
        //                                    fieldPromises.push(termsPromise);
        //                                    break;
        //                                case "multiMetadata":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //choices
        //                                    //displayValue
        //                                    //allowFillIn
        //                                    //default
        //                                    var termsPromise = getTermsForField(listName, columnName);
        //                                    termsPromise.then(function (terms) {
        //                                        setTermsForField(column, terms);
        //                                    }, function (reason) {

        //                                    });
        //                                    fieldPromises.push(termsPromise);
        //                                    break;
        //                                case "link":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //format (hyperlink/picture)
        //                                    break;
        //                                case "number":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //default
        //                                    //minmax
        //                                    //decimalplaces
        //                                    //showaspercentage
        //                                    break;
        //                                case "currency":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //default
        //                                    //minmax
        //                                    //decimalplaces
        //                                    //format
        //                                    break;
        //                                case "dateTime":
        //                                    fieldPromise = getListField(listName, columnName);
        //                                    fieldPromise.then(function (field) {
        //                                        console.log(listName + ":" + columnName);
        //                                        console.log(field);
        //                                    }, function (reason) {

        //                                    })
        //                                    fieldPromises.push(fieldPromise);
        //                                    //req
        //                                    //default
        //                                    //format (dateOnly/dateTime) (standard/friendly)
        //                                    break;
        //                                default:
        //                                    //throw error.
        //                                    break;
        //                            }
        //                        }
        //                        else {
        //                            //throw error.
        //                        }
        //                    }
        //                }
        //            }
        //            else {
        //                //throw error.
        //            }
        //        }
        //    }
        //    return $q.all(fieldPromises);
        //}

        function getListGuid(listName) {
            var deferred = $q.defer();

            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            var list = web.get_lists().getByTitle(listName);
            context.load(list)
            context.executeQueryAsync(success, failure);

            return deferred.promise;

            function success() {
                var listGuid = list.get_id();
                deferred.resolve(listGuid.toString());
            }

            function failure(sender, args) {
                deferred.reject(sender, args);
            }
        }

        function getListField(listName, fieldName) {
            var fieldDeferred = $q.defer();
            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            var list = web.get_lists().getByTitle(listName);
            var fields = list.get_fields();
            var field = fields.getByInternalNameOrTitle(fieldName);
            context.load(field);
            context.executeQueryAsync(onFieldSucceeded, function (sender, args) {
                console.log(fieldName);
                console.log(args.get_message());
            });
            
            return fieldDeferred.promise;

            function onFieldSucceeded() {
                fieldDeferred.resolve(field);
            }
        }

        function parseMetadataSchema(column) {
            var xmlDoc = $.parseXML(column.SchemaXml);
            var xml = $(xmlDoc);
            var field = xml.find("Field");
            var type = "";
            if (field !== undefined) {
                type = field.attr("Type");
            }
            var properties = xml.find("Property");
            for (var i = 0; i < properties.length; i++) {
                var propertyName = properties[i].firstChild.textContent === undefined ? properties[i].firstChild.text : properties[i].firstChild.textContent;
                var propertyValue = properties[i].lastChild.textContent === undefined ? properties[i].lastChild.text : properties[i].lastChild.textContent;
                if (propertyName !== propertyValue) {
                    switch (propertyName) {
                        case "SspId":
                            column.SspId = propertyValue;
                            break;
                        case "GroupId":
                            column.GroupId = propertyValue;
                            break;
                        case "TermSetId":
                            column.TermSetId = propertyValue;
                    }
                }
            }
            return type;
        }

        function setItemValues(columns, listItem, itemProps) {
            for (var columnName in columns) {
                if (columns.hasOwnProperty(columnName) && itemProps[columnName] !== undefined && itemProps[columnName].display && columns[columnName].type !== undefined && columns[columnName].inputType != undefined) {
                    var columnType = columns[columnName].type;
                    var inputType = columns[columnName].inputType;
                    var formValue = itemProps[columnName].value;
                    var columnValue;
                    switch (columnType) {
                        case "text":
                            columnValue = processText(formValue, inputType);
                            break;
                        case "multiText":
                            columnValue = processMultiText(formValue, inputType);
                            break;
                        case "choice":
                            columnValue = processChoice(formValue, inputType);
                            break;
                        case "multiChoice":
                            columnValue = processMultiText(formValue, inputType);
                            break;
                        case "lookup":
                            columnValue = processLookup(formValue, inputType);
                            break;
                        case "multiLookup":
                            columnValue = processMultiLookup(formValue, inputType);
                            break;
                        case "yesNo":
                            columnValue = processYesNo(formValue, inputType);
                            break;
                        case "person":
                            columnValue = processPerson(formValue, inputType);
                            break;
                        case "multiPerson":
                            columnValue = processMultiPerson(formValue, inputType);
                            break;
                        case "metadata":
                            columnValue = processMetadata(formValue, inputType);
                            break;
                        case "multiMetadata":
                            columnValue = processMultiMetadata(formValue, inputType);
                            break;
                        case "link":
                            columnValue = processLink(formValue, inputType);
                            break;
                        case "number":
                            columnValue = processNumber(formValue, inputType);
                            break;
                        case "currency":
                            columnValue = processCurrency(formValue, inputType);
                            break;
                        case "dateTime":
                            columnValue = processDateTime(formValue, inputType);
                            break;
                        default:
                            columnValue = null;
                            break;
                    }
                    listItem.set_item(columnName, columnValue);
                }
            }
            return listItem;
        }

        /**/

        return formSubmitPromise;

        function processText(formValue, inputType) {
            if (formValue !== undefined && formValue !== "") {
                return formValue;
            }
            else {
                return null;
            }
        }

        function processMultiText(formValue, inputType) {
            if (formValue !== undefined && formValue !== "" && formValue !== "<br>") {
                return formValue;
            }
            else {
                return null;
            }
        }

        function processChoice(formValue, inputType) {
            if (formValue !== undefined && formValue !== "") {
                return formValue;
            }
            else {
                return null;
            }
        }

        function processMetadata(formValue, inputType) {
            if (formValue !== undefined && formValue !== "") {
                var metadataItem = metadataItems[formValue];
                return metadataItem.label + "|" + formValue;
            }
            else {
                return null;
            }
        }

        function processMetadataMulti(formValue, inputType) {
            if (formValue !== undefined) {
                var metadataValues = [];
                for (var key in formValue) {
                    if (fieldValues[key]) {
                        metadataValues.push("-1");
                        metadataValues.push(metadataItems[key].label + "|" + key);
                    }
                }
                return metadataValues.join(";#");
            }
            else {
                return null;
            }
        }

        function processDate(formValue, inputType) {
            if (formValue !== undefined && formValue !== "") {
                var workingDate = new Date(formValue);
                return workingDate.toISOString();
            }
            else {
                return null;
            }
        }

        function processYesNo(formValue, inputType) {
            if (formValue !== undefined && formValue !== "") {
                return formValue;
            }
            else {
                return null;
            }
        }

        function processPeopePicker(formValue, inputType) {
            var userInfoArray = programDetails[field];
            if (userInfoArray !== undefined && (visibleElements[field] || visibleElements[field] === undefined)) {
                if (userInfoArray.length > 0) {
                    var userInfo = userInfoArray[0]
                    var user = SP.FieldUserValue.fromUser(userInfo.Key);
                    //console.log(userInfo);
                    return { fieldType: "user", value: user };
                }
            }
            return null;
        }

        function processPeoplePickerMulti(formValue, inputType) {
            var userInfoArray = programDetails[field];
            if (userInfoArray !== undefined && (visibleElements[field] || visibleElements[field] === undefined)) {
                var userArray = [];
                for (var i = 0; i < userInfoArray.length; i++) {
                    var userInfo = userInfoArray[i]
                    var user = SP.FieldUserValue.fromUser(userInfo.Key);
                    userArray.push(user);
                }
                if (userArray.length > 0) {
                    return { fieldType: "user", value: userArray };
                }
            }
            return null;
        }

        /**/
    }
})(angular, jQuery)