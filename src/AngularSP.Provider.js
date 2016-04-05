var app = app || angular.module("angularSP", ['ui.bootstrap', 'ui.grid', 'ui.grid.selection']);

(function (ng, $) {

    app.provider("spFormProvider", spFormProvider);

    function spFormProvider() {
        var lists = {};

        this.defineLists = function (listDefinitions) {
            lists = listDefinitions;
        }

        this.$get = ['$q', '$log', spFormFactory];

        function spFormFactory($q, $log) {
            return new SPForm($q, $log, lists);
        }
    }

    function SPForm($q, $log, lists) {
        var ignoreFields = ['AppAuthor', 'AppEditor', 'FolderChildCount', 'ItemChildCount', 'UniqueId', 'SyncClientId', 'SortBehavior', 'ScopeId', 'ProgId', 'MetaInfo', 'Last_x0020_Modified', 'FileDirRef', 'FSObjType', 'Created_x0020_Date', 'FileRef'];

        var listsInformation = null;

        listsInformationDeferred = $q.defer();

        this.listsInformationPromise = listsInformationDeferred.promise;

        initializeLists();

        this.createListItem = createListItem;
        this.createListItems = createListItems;
        this.getListItem = getListItem;
        this.getListItems = getListItems;
        this.updateListItem = updateListItem;
        this.deleteListItem = deleteListItem;
        this.commitListItems = commitListItems;


        function initializeLists() {

            readySharePoint(sharePointReady);

            function sharePointReady() {
                var listsArray = [];
                for (var listName in lists) {
                    if (lists.hasOwnProperty(listName)) {
                        var list = lists[listName];
                        setDisplayName(listName, list);
                        listsArray.push(list);
                    }
                }
                loadListInformation(listsArray);

                function setDisplayName(listName, list) {
                    if (list.DisplayName === undefined) {
                        list.DisplayName = listName;
                    }
                }
            }
        }

        function createListItem(listName, itemProperties) {
            var createDeferred = $q.defer();
            var listInfo = null;
            try {
                listInfo = getListInformation(listName);
            }
            catch (e) {
                createDeferred.reject(e.message);
                return;
            }
            readySharePoint(sharePointReady);
            return createDeferred.promise;

            function sharePointReady() {
                var context = getContext(listInfo);
                var web = context.get_web();
                var spList = web.get_lists().getByTitle(listInfo.DisplayName);
                var itemCreationInfo = new SP.ListItemCreationInformation();
                var spListItem = spList.addItem(itemCreationInfo);
                
                try {
                    setListItemValues(spListItem, listInfo.Columns, itemProperties);
                }
                catch (e) {
                    createDeferred.reject(e.message);
                    return;
                }
                
                spListItem.update();
                context.load(spListItem);
                context.executeQueryAsync(function () {
                    createDeferred.resolve(spListItem.get_id());
                }, function (sender, args) {
                    createDeferred.reject(args.get_message());
                });
            }
        }

        function createListItems() {
            $log.error("Function not implemented.");
        }

        function getListItem(listName, itemId) {
            var listItemDeferred = $q.defer();
            
            var listInfo = null;
            try {
                listInfo = getListInformation(listName);
            }
            catch (e) {
                listItemDeferred.reject(e.message);
                return listItemDeferred.promise;
            }
            readySharePoint(sharePointReady);

            return listItemDeferred.promise;

            function sharePointReady() {
                var context = getContext(listInfo);
                var spWeb = context.get_web();
                var spList = spWeb.get_lists().getByTitle(listInfo.DisplayName);
                var spItem = spList.getItemById(itemId);
                context.load(spItem);
                context.executeQueryAsync(getListItemSuccess, getListItemFailure);
                
                function getListItemSuccess() {
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
                            }
                        }
                    }
                    inputValues.Id = spItemValues.ID;
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
                }
                function getListItemFailure(sender, args) {
                    listItemDeferred.reject(args.get_message());
                }
            }
        }

        function getListItems() {
            $log.error("Function not implemented.");
        }

        function updateListItem(listName, itemProperties, overrideConflict) {
            var updateDeferred = $q.defer();

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

            return updateDeferred.promise;

            function sharePointReady() {
                var listName = listInfo.DisplayName;

                var context = SP.ClientContext.get_current();
                var web = context.get_web();
                var spList = web.get_lists().getByTitle(listName);
                var itemCreationInfo = new SP.ListItemCreationInformation();
                var listItem = spList.getItemById(itemProps.Id);

                if (overrideConflict) {
                    context.load(listItem);
                    context.executeQueryAsync(updateItemValues, function (sender, args) {
                        updateDeferred.reject(args.get_message());
                    });
                }
                else {
                    updateItemValues();
                }

                function updateItemValues() {
                    var columns = listInfo.Columns;
                    for (var columnName in columns) {
                        if (columns.hasOwnProperty(columnName)) {
                            var columnInfo = columns[columnName];
                            if (itemProps[columnName] !== undefined) {
                                var inputValue = itemProps[columnName];
                                var itemValue = getListItemValue(columnInfo, inputValue);
                                listItem.set_item(columnName, itemValue);
                            }
                        }
                    }
                    listItem.update();
                    context.load(listItem);
                    context.executeQueryAsync(function () {
                        updateDeferred.resolve(listItem.get_id());
                    }, function (sender, args) {
                        updateDeferred.reject(args.get_message());
                    });
                }
            }
        }

        function deleteListItem(listName, itemId) {
            var deleteDeferred = $q.defer();

            var listDisplayName = listsInfo[listName].DisplayName;

            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            var list = web.get_lists().getByTitle(listDisplayName);
            var item = list.getItemById(itemId);
            item.deleteObject();
            context.executeQueryAsync(success, failure);

            return deleteDeferred.promise;

            function success() {
                deleteDeferred.resolve();
            }

            function failure(sender, args) {
                deleteDeferred.reject(args.get_message());
            }
        }

        function commitListItems() {
            $log.error("Function not implemented.");
        }

        /*Private*/

        function AngularSPException(errorMessage) {
            this.message = errorMessage;
            this.toString = function () {
                return this.message;
            };
        }

        function getListInformation(listName) {
            if (listsInformation !== null && listsInformation[listName] !== undefined) {
                return listsInformation[listName];
            }
            else {
                throw new AngularSPException("List information not initialized.  Either there was an initialization error or you need to wait until the listsInformationPromise is resolved.");
            }
        }

        function loadListInformation(listsArray) {
            var context = null;
            var spFields = null;
            var list = listsArray.pop();
            if (list !== undefined) {
                context = getContext(list);
                var spWeb = context.get_web();
                var spLists = spWeb.get_lists();
                var spList = spLists.getByTitle(list.DisplayName);
                spFields = spList.get_fields();
                context.load(spFields);
                context.executeQueryAsync(getFieldsSuccess, getFieldsFailure);
            }
            else {
                listsInformation = lists;
                listsInformationDeferred.resolve(lists);
            }

            function getFieldsSuccess() {
                list.Columns = {};
                var fieldEnum = spFields.getEnumerator();
                while (fieldEnum.moveNext()) {
                    var spField = fieldEnum.get_current();
                    var internalName = spField.get_internalName();
                    if (ignoreFields.indexOf(internalName) < 0) {
                        list.Columns[internalName] = {
                            Field: spField
                        };
                        context.load(spField);
                    }
                }
                context.executeQueryAsync(getFieldInformationSuccess, getFieldInformationFailure);

                function getFieldInformationSuccess() {
                    var hasLookups = false;
                    for (var columnName in list.Columns) {
                        if (list.Columns.hasOwnProperty(columnName)) {
                            var column = list.Columns[columnName];
                            hasLookups = setColumnInfo(context, column) || hasLookups;
                        }
                    }
                    if (hasLookups === true) {
                        context.executeQueryAsync(getLookupsSuccess, getLookupsFailure);
                    }
                    else {
                        loadListInformation(listsArray);
                    }

                    function getLookupsSuccess() {
                        setLookupInfo(list.Columns);
                        loadListInformation(listsArray);
                    }
                    function getLookupsFailure(sender, args) {
                        listsInformationDeferred.reject(args.get_message());
                    }
                }

                function getFieldInformationFailure(sender, args) {
                    listsInformationDeferred.reject(args.get_message());
                }
            }

            function getFieldsFailure(sender, args) {
                listsInformationDeferred.reject(args.get_message());
            }
        }

        function readySharePoint(callback) {
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                SP.SOD.executeFunc('sp.taxonomy.js', 'SP.Taxonomy.TaxonomySession', function () {
                    SP.SOD.executeFunc('SP.UserProfiles.js', 'SP.UserProfiles.PeopleManager', callback);
                });
            });
        }

        function getContext(list) {
            if (list.ContextUrl === undefined) {
                return SP.ClientContext.get_current();
            }
            else {
                return new SP.ClientContext(list.ContextUrl);
            }
        }

        function setColumnInfo(context, column) {
            var hasLookups = false;
            var spField = column.Field;
            column.Title = spField.get_title();
            column.Id = spField.get_id().toString();
            column.IsRequired = spField.get_required();
            column.SchemaXml = spField.get_schemaXml();
            column.DefaultValue = spField.get_defaultValue();
            column.IsReadOnly = spField.get_readOnlyField();
            column.Scope = spField.get_scope();
            column.FieldType = spField.get_fieldTypeKind();
            switch (column.FieldType) {
                case 0: //metadata field
                    var taxFieldType = parseMetadataSchema(column);
                    column.AllowMultipleValues = (taxFieldType === 'TaxonomyFieldTypeMulti');
                    column.RawTerms = loadMetadataTerms(context, column);
                    column.SetLookupFunction = setMetadataTerms;
                    hasLookups = true;
                    break;
                case 2: //single text field
                    break;
                case 3: //multi text field
                    var multiLineTextField = context.castTo(spField, SP.FieldMultiLineText);
                    column.IsRichText = multiLineTextField.get_richText();
                    break;
                case 4: //datetime field
                    break;
                case 6: //choice field
                    var choiceField = context.castTo(spField, SP.FieldChoice);
                    column.Choices = choiceField.get_choices();
                    column.FillInChoice = choiceField.get_fillInChoice();
                    column.EditFormat = choiceField.get_editFormat();
                    break;
                case 7: //lookup field
                    var lookupField = context.castTo(spField, SP.FieldLookup);
                    column.AllowMultipleValues = lookupField.get_allowMultipleValues();
                    column.LookupField = lookupField.get_lookupField();
                    column.LookupList = lookupField.get_lookupList();
                    column.RawLookupItems = loadLookupOptions(context, column);
                    column.SetLookupFunction = setLookupOptions;
                    hasLookups = true;
                    break;
                case 8: //boolean field
                    break;
                case 9: //number field
                    break;
                case 10: //currency field
                    break;
                case 11: //link field
                    break;
                case 15: //multiple choice field
                    var multiChoiceField = context.castTo(spField, SP.FieldMultiChoice);
                    column.Choices = multiChoiceField.get_choices();
                    column.FillInChoice = multiChoiceField.get_fillInChoice();
                    break;
                case 20: //person field
                    var userField = context.castTo(spField, SP.FieldUser);
                    column.SelectionGroup = userField.get_selectionGroup();
                    column.PeopleOnly = (userField.get_selectionMode() === 0);
                    column.AllowMultipleValues = userField.get_allowMultipleValues();
                    break;
                default:
                    //$log.warn('Field \"' + column.Title + '\" has type \"' + column.FieldType + '\" which is not supported.')
                    break;
            }
            return hasLookups;
        }

        function setLookupInfo(columns) {
            for (var columnName in columns) {
                if (columns.hasOwnProperty(columnName)) {
                    var column = columns[columnName];
                    if (column.SetLookupFunction !== undefined) {
                        column.SetLookupFunction(column);
                    }
                }
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

        function loadMetadataTerms(context, column) {
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
            var termStores = taxSession.get_termStores();
            var termStore = termStores.getById(column.SspId);
            var termSet = termStore.getTermSet(column.TermSetId);
            var rawTerms = termSet.getAllTerms();
            context.load(rawTerms);
            return rawTerms;
        }

        function loadLookupOptions(context, column) {
            var query = new SP.CamlQuery();
            var web = context.get_web();
            var list = web.get_lists().getById(column.LookupList);
            var rawLookupItems = list.getItems(query);
            context.load(rawLookupItems);
            return rawLookupItems;
        }

        function setMetadataTerms(column) {
            column.Terms = [];
            column.TermsById = {};
            column.TermsByLabel = {};
            var termEnum = column.RawTerms.getEnumerator();
            while (termEnum.moveNext()) {
                var currentTerm = termEnum.get_current();
                var currentTermId = currentTerm.get_id().toString();
                var currentTermLabel = currentTerm.get_name();
                var termObj = { 'Id': currentTermId, 'Label': currentTermLabel }
                column.Terms.push(termObj);
                column.TermsById[currentTermId] = termObj;
                column.TermsByLabel[currentTermLabel] = termObj;
            }
        }

        function setLookupOptions(column) {
            column.LookupItems = [];
            column.LookupItemsById = {};
            column.LookupItemsByLabel = {};
            var lookupItemsEnum = column.RawLookupItems.getEnumerator();
            while (lookupItemsEnum.moveNext()) {
                var lookupItem = lookupItemsEnum.get_current();
                var lookupItemValues = lookupItem.get_fieldValues();
                var lookupItemId = lookupItemValues.ID;
                var lookupItemLabel = lookupItemValues[column.LookupField];
                var lookupItemObj = { 'Id': lookupItemId, 'Label': lookupItemLabel };
                column.LookupItems.push(lookupItemObj);
                column.LookupItemsById[lookupItemId] = lookupItemObj;
                column.LookupItemsByLabel[lookupItemLabel] = lookupItemObj;
            }
        }

        function setListItemValues(spListItem, columns, values) {
            //for (var columnName in listInfo.Columns) {
            //    if (listInfo.Columns.hasOwnProperty(columnName)) {
            //        var columnInfo = listInfo.Columns[columnName];
            //        if (itemProperties[columnName] !== undefined) {
            //            var inputValue = itemProperties[columnName];
            //            var itemValue = getListItemValue(columnInfo, inputValue);
            //            spListItem.set_item(columnName, itemValue);
            //        }
            //    }
            //}
            for (var columnName in values) {
                if (values.hasOwnProperty(columnName)) {
                    if (columns[columnName] !== undefined) {
                        var columnInfo = columns[columnName];
                        if (columnInfo.IsReadOnly === false) {
                            var inputValue = values[columnName];
                            var itemValue = getListItemValue(columnInfo, inputValue);
                            spListItem.set_item(columnName, itemValue);
                        }
                        else {
                            throw new AngularSPException("Trying to write to read only field: \'" + columnName + "\'");
                        }
                    }
                    else {
                        throw new AngularSPException("There is a mismatch between the Item Properties and the List Column Information for field: \'" + columnName + "\'");
                    }
                }
            }
        }

        function getListItemValue(column, inputValue) {
            var returnValue = null;
            switch (column.FieldType) {
                case 0: //metadata field
                    if (column.AllowMultipleValues) {
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
                        returnValue = metadataValues.join(";#");
                    }
                    else {
                        var metadataItem = column.TermsById[inputValue];
                        returnValue = metadataItem.label + "|" + metadataItem.id;
                    }
                    break;
                case 2: //single text field
                    returnValue = inputValue;
                    break;
                case 3: //multi text field
                    returnValue = inputValue;
                    break;
                case 4: //datetime field
                    returnValue = inputValue.toISOString();
                    break;
                case 6: //choice field
                    returnValue = inputValue;
                    break;
                case 7: //lookup field
                    if (column.AllowMultipleValues) {
                        returnValue = [];
                        for (var value in inputValue) {
                            if (inputValue.hasOwnProperty(value)) {
                                if (inputValue[value]) {
                                    var lookupValue = new SP.FieldLookupValue();
                                    lookupValue.set_lookupId(value);
                                    returnValue.push(lookupValue);
                                }
                            }
                        }
                    }
                    else {
                        returnValue = new SP.FieldLookupValue();
                        returnValue.set_lookupId(inputValue);
                    }
                    break;
                case 8: //boolean field
                    if (inputValue) {
                        returnValue = 1;
                    }
                    else {
                        returnValue = 0;
                    }
                    break;
                case 9: //number field
                    returnValue = inputValue;
                    break;
                case 10: //currency field
                    returnValue = inputValue;
                    break;
                case 11: //link field
                    returnValue = inputValue;
                    break;
                case 15: //multiple choice field
                    returnValue = [];
                    for (var value in inputValue) {
                        if (inputValue.hasOwnProperty(value)) {
                            if (inputValue[value]) {
                                itemValue.push(value);
                            }
                        }
                    }
                    break;
                case 20: //person field
                    if (column.AllowMultipleValues) {
                        returnValue = [];
                        for (var i = 0; i < inputValue.length; i++) {
                            var person = inputValue[i];
                            returnValue.push(SP.FieldUserValue.fromUser(person.Key));
                        }
                    }
                    else {
                        if (inputValue.length > 0) {
                            var person = inputValue[0];
                            returnValue = SP.FieldUserValue.fromUser(person.Key);
                        }
                    }
                    break;
                default:
                    //$log.warn('Field \"' + column.Title + '\" has type \"' + column.FieldType + '\" which is not supported.')
                    break;
            }
            return returnValue;
        }
    }

})(angular, jQuery)