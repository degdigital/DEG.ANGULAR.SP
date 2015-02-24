'use strict';

(function (ng) {

    ng.module("angularSP", [])
        .factory("spListFactory", ['$q', spListFactory]);

    function spListFactory($q) {

        var service = {
            initFactory: initFactory,
            createListItem: createListItem,
            createListItems: createListItems,
            getListItem: getListItem,
            getListItems: getListItems,
            updateListItem: updateListItem,
            deleteListItem: deleteListItem
        }

        var serverRelativeUrl;

        var columnDefinitions;

        var listDefinitions;

        var columnsInfo;

        var listsInfo;

        return service;

        function initFactory(columnDefs, listDefs) {

            columnDefinitions = columnDefs;
            listDefinitions = listDefs;

            columnsInfo = ng.copy(columnDefs);
            listsInfo = ng.copy(listDefs);

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

                var fieldData = loadColumns(columnDefs);

                context.executeQueryAsync(success, failure);

                function success() {
                    serverRelativeUrl = web.get_serverRelativeUrl();
                    setColumnInfo(columnsInfo, fieldData);
                    console.log(columnsInfo);
                    initDeferred.resolve();
                }

                function failure(sender, args) {
                    initDeferred.reject(args.get_message());
                }
            }
        }

        function createListItem(listName, itemProps) {
            var createDeferred = $q.defer();

            if (listDefinitions !== undefined) {
                if (listDefinitions[listName] !== undefined) {
                    var listDef = listDefinitions[listName];
                    var columns = listDef.Columns;

                }
            }

            return createDeferred.promise;
        }

        function createListItems(listName, itemsProps) {
            var createDeferred = $q.defer();



            return createDeferred.promise;
        }

        function getListItem(listName, ItemId) {

        }

        function getListItems(listName, params) {

        }

        function updateListItem(listName, itemProps) {

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

        function loadColumns(columnDefs) {
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

        function setColumnInfo(columnsInfo, fieldData) {
            var context = new SP.ClientContext.get_current();
            for (var i = 0; i < fieldData.length; i++) {
                var field = fieldData[i];
                var internalName = field.get_internalName();
                if (columnsInfo[internalName] !== undefined) {
                    var column = columnsInfo[internalName];
                    column.IsRequired = field.get_required();
                    column.SchemaXml = field.get_schemaXml();
                    column.DefaultValue = field.get_defaultValue();
                    switch (column.Type) {
                        case "text":
                            break;
                        case "multiText":
                            break;
                        case "choice":
                            var choiceField = context.castTo(field, SP.FieldChoice);
                            //req
                            //choices
                            //default
                            //displayAs
                            break;
                        case "multiChoice":
                            var choiceField = context.castTo(field, SP.FieldChoice);
                            //req
                            //choices
                            //default
                            break;
                        case "lookup":
                            //req
                            //choices
                            break;
                        case "multiLookup":
                            //req
                            //choices
                            break;
                        case "yesNo":
                            break;
                        case "person":
                            //req
                            //peopleOnly/peopleAndGroups
                            //chooseFrom(group)
                            //showField
                            break;
                        case "multiPerson":
                            //req
                            //peopleOnly/peopleAndGroups
                            //chooseFrom(group)
                            //showField
                            break;
                        case "metadata":
                            //req
                            //choices
                            //displayValue
                            //allowFillIn
                            //default
                            var termsPromise = getTermsForField(listName, columnName);
                            termsPromise.then(function (terms) {
                                setTermsForField(column, terms);
                            }, function (reason) {

                            });
                            fieldPromises.push(termsPromise);
                            break;
                        case "multiMetadata":
                            //req
                            //choices
                            //displayValue
                            //allowFillIn
                            //default
                            var termsPromise = getTermsForField(listName, columnName);
                            termsPromise.then(function (terms) {
                                setTermsForField(column, terms);
                            }, function (reason) {

                            });
                            fieldPromises.push(termsPromise);
                            break;
                        case "link":
                            //req
                            //format (hyperlink/picture)
                            break;
                        case "number":
                            //req
                            //default
                            //minmax
                            //decimalplaces
                            //showaspercentage
                            break;
                        case "currency":
                            //req
                            //default
                            //minmax
                            //decimalplaces
                            //format
                            break;
                        case "dateTime":
                            //req
                            //default
                            //format (dateOnly/dateTime) (standard/friendly)
                            break;
                        default:
                            //throw error.
                            break;
                    }
                }
                else {
                    //Very strange error.
                    console.log("Field data mismatch.");
                }
            }
        }

        function loadListInfoProperties(listsInfo) {
            var fields = [];
            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            for (var listName in listsInfo) {
                if (listsInfo.hasOwnProperty(listName)) {
                    var listDef = listsInfo[listName];
                    var listTitle = listName;
                    if (list.DisplayName !== undefined) {
                        listTitle = listDef.DisplayName;
                    }
                    var list = web.get_lists().getByTitle(listTitle);
                    if (list.Columns !== undefined) {
                        loadListFields(context, list, list.Columns, fields);
                    }
                    else {
                        //Report missing column defenitions Error.
                    }
                }
            }
            return fields;
        }

        function loadListFields(context, list, listColumns, fields) {
            var fields = list.get_fields();
            for (var columnName in listColumns) {
                if (listColumns.hasOwnProperty(columnName)) {
                    var field = fields.getByInternalNameOrTitle(columnName);
                    context.load(field);
                    fields.push(field);
                }
            }
        }

        function setListInfoProperties(fields) {
            for (var i = 0; i < fields.length; i++) {
                var field = fields[i];
                var internalName = field.get_internalName();
            }
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

        function getTermsForField(listName, fieldName) {
            var termsDeferred = $q.defer();
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

            return termsDeferred.promise;

            function onFieldSucceeded() {
                var fieldSchema = field.get_schemaXml();
                var sspInfo = parseMetadataSchema(fieldSchema);
                var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
                var termStores = taxSession.get_termStores();
                var termStore = termStores.getById(sspInfo.sspId);
                var termSet = termStore.getTermSet(sspInfo.termSetId);
                var terms = termSet.getAllTerms();
                context.load(terms);
                context.executeQueryAsync(onTaxSucceeded, function (sender, args) {
                    console.log(fieldName);
                    console.log(args.get_message());
                });

                function onTaxSucceeded() {
                    termsDeferred.resolve(terms);
                }
            }
        }

        function setTermsForField(column, terms) {
            var termEnum = terms.getEnumerator();
            column.Terms = [];
            column.TermsById = {};
            column.TermsByLabel = {};
            while (termEnum.moveNext()) {
                var currentTerm = termEnum.get_current();
                var currentTermId = currentTerm.get_id().toString();
                var currentTermLabel = currentTerm.get_name();
                var termObj = { id: currentTermId, label: currentTermLabel }
                column.Terms.push(termObj);
                column.TermsById[currentTermId] = termObj;
                column.TermsByLabel[currentTermLabel] = termObj;
            }
            callbackFunction(termInfo);
        }

        function parseMetadataSchema(schema) {
            var xmlDoc = $.parseXML(schema);
            var xml = $(xmlDoc);
            var sspInfo = { sspId: "", groupId: "", termSetId: "" };
            var properties = xml.find("Property");
            for (i = 0; i < properties.length; i++) {
                var propertyName = properties[i].firstChild.textContent === undefined ? properties[i].firstChild.text : properties[i].firstChild.textContent;
                var propertyValue = properties[i].lastChild.textContent === undefined ? properties[i].lastChild.text : properties[i].lastChild.textContent;
                switch (propertyName) {
                    case "SspId":
                        sspInfo.sspId = propertyValue;
                        break;
                    case "GroupId":
                        sspInfo.groupId = propertyValue;
                        break;
                    case "TermSetId":
                        sspInfo.termSetId = propertyValue;
                }
            }
            return sspInfo;
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
})(angular)