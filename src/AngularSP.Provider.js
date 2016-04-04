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
        //$log.log(lists);

        this.listInformationPromise = initializeLists();

        this.createListItem = createListItem;
        this.createListItems = createListItems;
        this.getListItem = getListItem;
        this.getListItems = getListItems;
        this.updateListItem = updateListItem;
        this.deleteListItem = deleteListItem;
        this.commitListItems = commitListItems;


        function initializeLists() {

            var initDeferred = $q.defer();

            readySharePoint(sharePointReady);

            return initDeferred.promise;

            function sharePointReady() {
                var listsArray = [];
                for (var listName in lists) {
                    if (lists.hasOwnProperty(listName)) {
                        var list = lists[listName];
                        setDisplayName(listName, list);
                        listsArray.push(list);
                    }
                }
                var listInfoPromise = loadListInformation(listsArray);

                function setDisplayName(listName, list) {
                    if (list.DisplayName === undefined) {
                        list.DisplayName = listName;
                    }
                }

                function loadListInformation(listsArray) {
                    var context;
                    var list = listsArray.pop();
                    if (list !== undefined) {
                        var context = getContext(list);
                        var spWeb = context.get_web();
                        var spLists = spWeb.get_lists();
                        var spList = spLists.getByTitle(list.DisplayName);
                        var spFields = spList.get_fields();
                        context.load(spFields);
                        context.executeQueryAsync(getFieldsSuccess, getFieldsFailure);
                    }
                    else {
                        initDeferred.resolve(lists);
                    }
                    
                    function getFieldsSuccess() {
                        list.Columns = {};
                        var hasLookups = false;
                        var fieldEnum = spFields.getEnumerator();
                        while (fieldEnum.moveNext()) {
                            var spField = fieldEnum.get_current();
                            hasLookups = setColumnInfo(context, list.Columns, spField);
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
                            initDeferred.reject(args.get_message());
                        }
                    }

                    function getFieldsFailure(sender, args) {
                        initDeferred.reject(args.get_message());
                    }
                }
            }
        }

        function createListItem() {

        }

        function createListItems() {

        }

        function getListItem() {

        }

        function getListItems() {

        }

        function updateListItem() {

        }

        function deleteListItem() {

        }

        function commitListItems() {

        }

        /*Private*/

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

        function setColumnInfo(context, columns, spField) {
            var hasLookups = false;
            var internalName = spField.get_internalName();
            var column = columns[internalName] = {};
            column.Title = spField.get_title();
            column.Id = spField.get_id().toString();
            column.IsRequired = spField.get_required();
            column.SchemaXml = spField.get_schemaXml();
            column.DefaultValue = spField.get_defaultValue();
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
                    break
                case 5: //choice field
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
                    $log.warn('Field ' + column.Title + ' has type ' + column.FieldType + ' which is not supported.')
                    break;
            }
            return hasLookups;
        }

        function setLookupInfo(columns) {
            for (var columnName in columns) {
                if (lists.hasOwnProperty(columnName)) {
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
    }

})(angular, jQuery)