var app = app || angular.module("angularSP", ['ui.bootstrap', 'ui.grid', 'ui.grid.selection']);

(function (ng, $) {

    app.directive("contenteditable", function () {
            return {
                restrict: "A",
                require: "?ngModel",
                link: function (scope, element, attrs, ngModel) {
                    if (!ngModel) {
                        return;
                    }

                    function read() {
                        var html = element.html();
                        if (html === "<br>") {
                            ngModel.$setViewValue("");
                        }
                        else {
                            ngModel.$setViewValue(html);
                        }
                        
                    }

                    ngModel.$render = function () {
                        element.html(ngModel.$viewValue || "");
                    };

                    element.bind("blur keyup change", function () {
                        scope.$apply(read);
                    });
                }
            };
        });

    app.directive("ngspColumn", angularSpField);

    app.directive("ngspList", angularSpList);
    
    function angularSpField ($compile) {
        return {
            restrict: "A",
            require: "ngModel",
            scope: {
                ngspColumn: '=',
                ngspReadOnly: '='
            },
            link: link
        }

        function link(scope, element, attrs, ngModel) {
            scope.ngspColumn = {};
            scope.$watch('ngspColumn', function () {
                if (scope.ngspColumn.Title !== undefined) {
                    element.html("");
                    element.append(generateTitleElement(scope.ngspColumn));
                    if (scope.ngspReadOnly !== undefined && scope.ngspReadOnly) {
                        element.append(generateDisplayElement(ngModel, scope));
                    }
                    else{
                        element.append(generateInputElement(element, ngModel, scope));
                    }
                }
            });
        }

        function generateTitleElement(column) {
            var titleElement = $("<label></label>");
            titleElement.text(column.Title);
            return titleElement;
        }

        function generateInputElement(element, ngModel, scope) {
            switch (scope.ngspColumn.InputType) {
                case "text":
                    return generateTextElement(ngModel, scope);
                    break;
                case "textBox":
                    return generateTextBoxElement(ngModel, scope);
                    break;
                case "nicEdit":
                    return generateNicEditElement(ngModel, scope);
                    break;
                case "dropDown":
                    return generateDropDownElement(ngModel, scope);
                    break;
                case "radio":
                    return generateRadioElement(ngModel, scope);
                    break;
                case "checkbox":
                    return generateCheckboxElement(ngModel, scope);
                    break;
                case "dateTime":
                    return generateDateTimeElement(ngModel, scope);
                    break;
                case "peoplePicker":
                    return generatePeoplePickerElement(element, ngModel, scope);
                    break;
                default:
                    console.log(scope.ngspColumn.InputType + " is currently unsupported as an input type.");
                    break;
            }
        }

        function generateTextElement(ngModel, scope) {
            //if (ngModel.$viewValue !== undefined) {
            //    scope.InputValue = ngModel.$viewValue;
            //}

            ngModel.$formatters.push(function (value) {
                scope.InputValue = value;
            });

            scope.textInputChanged = function () {
                ngModel.$setViewValue(scope.InputValue);
            }

            var inputElement = $("<input />");
            inputElement.attr("data-ng-model", "InputValue");
            inputElement.attr("data-ng-change", "textInputChanged()");
            if (scope.ngspColumn.Type === "text") {

            }
            else if (scope.ngspColumn.Type === "link") {

            }
            else if (scope.ngspColumn.Type === "number") {

            }
            else if (scope.ngspColumn.Type === "currency") {

            }
            else {
                console.log(scope.ngspColumn.Type + "is not supported as a column type for text input.");
                return;
            }
            $compile(inputElement)(scope);
            return inputElement;
        }

        function generateTextBoxElement(ngModel, scope) {
            //if (ngModel.$viewValue !== undefined) {
            //    scope.InputValue = ngModel.$viewValue;
            //}

            ngModel.$formatters.push(function (value) {
                scope.InputValue = value;
            });

            scope.textAreaChanged = function () {
                ngModel.$setViewValue(scope.InputValue);
            }

            var inputElement = $("<textarea></textarea>");
            inputElement.attr("data-ng-model", "InputValue");
            inputElement.attr("data-ng-change", "textAreaChanged()");
            if (scope.ngspColumn.Type === "multiText") {

            }
            else {
                console.log(scope.ngspColumn.Type + "is not supported as a column type for text box.");
                return;
            }
            $compile(inputElement)(scope);
            return inputElement;
        }

        function generateNicEditElement(ngModel, scope) {
            //if (ngModel.$viewValue !== undefined) {
            //    scope.InputValue = ngModel.$viewValue;
            //}

            ngModel.$formatters.push(function (value) {
                scope.InputValue = value;
            });

            var outerElement = $("<span></span>");
            var panelElement = $("<div></div>");
            outerElement.append(panelElement);
            var contentElement = $("<div></div>");
            contentElement.attr("contenteditable", "true");
            contentElement.attr("data-ng-model", "InputValue");

            outerElement.append(contentElement);


            //JORDAN: parameterize this.
            var nicEditorConstrutorOption = { iconsPath: '/AngularSP/lib/images/nicEditorIcons.gif' };
            var nicEditorInstance = new nicEditor(nicEditorConstrutorOption);
            nicEditorInstance.setPanel(panelElement[0]);
            nicEditorInstance.addInstance(contentElement[0]);

            if (scope.ngspColumn.Type === "multiText") {

            }
            else {
                console.log(scope.ngspColumn.Type + "is not supported as a column type for nicEdit.");
                return;
            }

            $compile(outerElement)(scope);

            scope.$watch("InputValue", function () {
                ngModel.$setViewValue(scope.InputValue);
            })

            return outerElement;
        }

        function generateDropDownElement(ngModel, scope) {
            ngModel.$formatters.push(function (value) {
                scope.InputValue = value;
            });

            scope.dropDownChanged = function () {
                ngModel.$setViewValue(scope.InputValue);
            }

            var inputElement = $("<select></select>");
            inputElement.attr("data-ng-model", "InputValue");
            inputElement.attr("data-ng-change", "dropDownChanged()");
            if (scope.ngspColumn.Type === "choice") {
                inputElement.attr("data-ng-options", "choice for choice in ngspColumn.Choices");
            }
            else if (scope.ngspColumn.Type === "lookup") {
                inputElement.attr("data-ng-options", "lookupItem.id as lookupItem.label for lookupItem in ngspColumn.LookupItems");
            }
            else if (scope.ngspColumn.Type === "metadata") {
                inputElement.attr("data-ng-options", "term.id as term.label for term in ngspColumn.Terms");
            }
            else {
                console.log(scope.ngspColumn.Type + "is not supported as a column type for dropdown menus.");
                return;
            }
            $compile(inputElement)(scope);
            return inputElement;
        }

        function generateCheckboxElement(ngModel, scope) {

            //if (ngModel.$viewValue !== undefined) {
            //    scope.InputValue = ngModel.$viewValue;
            //}
            //else {
            //    scope.InputValue = {};
            //}

            ngModel.$formatters.push(function (value) {
                scope.InputValue = value;
            });

            scope.checkboxChanged = function () {
                ngModel.$setViewValue(scope.InputValue);
            }

            var outerElement = $("<span></span>");
            if (scope.ngspColumn.Type === "multiMetadata" || scope.ngspColumn.Type === "multiLookup" || scope.ngspColumn.Type === "multiChoice") {
                var repeatElement = $("<span></span>");
                var inputElement = $("<input />");
                inputElement.attr("data-ng-change", "checkboxChanged()");
                var labelElement = $("<label></label>");
                inputElement.attr("type", "checkbox");
                if (scope.ngspColumn.Type === "multiMetadata")
                {
                    repeatElement.attr("data-ng-repeat", "term in ngspColumn.Terms");
                    inputElement.attr("data-ng-model", "InputValue[term.id]");
                    labelElement.text("{{term.label}}");
                }
                else if (scope.ngspColumn.Type === "multiLookup") {
                    repeatElement.attr("data-ng-repeat", "lookupItem in ngspColumn.LookupItems");
                    inputElement.attr("data-ng-model", "InputValue[lookupItem.id]");
                    labelElement.text("{{lookupItem.label}}");
                }
                else if (scope.ngspColumn.Type === "multiChoice") {
                    repeatElement.attr("data-ng-repeat", "choice in ngspColumn.Choices");
                    inputElement.attr("data-ng-model", "InputValue[choice]");
                    labelElement.text("{{choice}}");
                }
                repeatElement.append(inputElement);
                repeatElement.append(labelElement);
            }
            else if (scope.ngspColumn.Type === "yesNo") {
                var inputElement = $("<input />");
                inputElement.attr("type", "checkbox");
                inputElement.attr("data-ng-model", "InputValue");
                inputElement.attr("data-ng-change", "checkboxChanged()");
                outerElement.append(inputElement);
            }
            else {
                console.log(scope.ngspColumn.Type + "is not supported as a column type for checkboxes.");
                return;
            }
            outerElement.append(repeatElement);
            $compile(outerElement)(scope);
            return outerElement;
        }

        function generateRadioElement(ngModel, scope) {

            ngModel.$formatters.push(function (value) {
                scope.InputValue = value;
            });

            scope.radioChanged = function () {
                if (scope.ngspColumn.Type === "yesNo") {
                    if (scope.InputValue === "yes") {
                        ngModel.$setViewValue(true);
                    }
                    else {
                        ngModel.$setViewValue(false);
                    }
                }
                else {
                    ngModel.$setViewValue(scope.InputValue);
                }
            }

            var outerElement = $("<span></span>");

            var radioName = scope.ngspColumn.Title;

            if (scope.ngspColumn.Type === "yesNo") {
                var outerYesElement = $("<span></span>");
                var yesElement = $("<input></input>");
                yesElement.attr("type", "radio");
                yesElement.attr("data-ng-model", "InputValue");
                yesElement.attr("data-ng-change", "radioChanged()");
                yesElement.attr("name", radioName);
                yesElement.attr("value", "yes");
                outerYesElement.append(yesElement);
                var yesLabel = $("<label></label>");
                yesLabel.text("Yes");
                outerYesElement.append(yesLabel);
                var outerNoElement = $("<span></span>");
                var noElement = $("<input></input>");
                noElement.attr("type", "radio");
                noElement.attr("data-ng-model", "InputValue");
                noElement.attr("data-ng-change", "radioChanged()");
                noElement.attr("name", radioName);
                noElement.attr("value", "no");
                outerNoElement.append(noElement);
                var noLabel = $("<label></label>");
                noLabel.text("No");
                outerNoElement.append(noLabel);
                outerElement.append(outerYesElement);
                outerElement.append(outerNoElement);
            }
            else if (scope.ngspColumn.Type === "choice") {
                var outerChoiceElement = $("<span></span>");
                outerChoiceElement.attr("data-ng-repeat", "choice in ngspColumn.Choices");
                var inputElement = $("<input></input>");
                inputElement.attr("type", "radio");
                inputElement.attr("data-ng-change", "radioChanged()");
                inputElement.attr("data-ng-model", "$parent.InputValue");
                inputElement.attr("name", radioName);
                inputElement.attr("value", "{{choice}}");
                var labelElement = $("<label></label>");
                labelElement.text("{{choice}}");
                outerChoiceElement.append(inputElement);
                outerChoiceElement.append(labelElement);
                outerElement.append(outerChoiceElement);
            }
            else if (scope.ngspColumn.Type === "lookup") {
                var outerLookupElement = $("<span></span>");
                outerLookupElement.attr("data-ng-repeat", "lookupItem in ngspColumn.LookupItems");
                var inputElement = $("<input></input>");
                inputElement.attr("type", "radio");
                inputElement.attr("data-ng-change", "radioChanged()");
                inputElement.attr("data-ng-model", "$parent.InputValue");
                inputElement.attr("name", radioName);
                inputElement.attr("value", "{{lookupItem.id}}");
                var labelElement = $("<label></label>");
                labelElement.text("{{lookupItem.label}}");
                outerLookupElement.append(inputElement);
                outerLookupElement.append(labelElement);
                outerElement.append(outerLookupElement);
            }
            else if (scope.ngspColumn.Type === "metadata") {
                var outerMetadataElement = $("<span></span>");
                outerMetadataElement.attr("data-ng-repeat", "term in ngspColumn.Terms");
                var inputElement = $("<input></input>");
                inputElement.attr("type", "radio");
                inputElement.attr("data-ng-change", "radioChanged()");
                inputElement.attr("data-ng-model", "$parent.InputValue");
                inputElement.attr("name", radioName);
                inputElement.attr("value", "{{term.id}}");
                var labelElement = $("<label></label>");
                labelElement.text("{{term.label}}");
                outerMetadataElement.append(inputElement);
                outerMetadataElement.append(labelElement);
                outerElement.append(outerMetadataElement);
            }
            else {
                console.log(scope.ngspColumn.Type + "is not supported as a column type for radio input.");
                return;
            }

            $compile(outerElement)(scope);
            return outerElement;
        }

        function generateDateTimeElement(ngModel, scope) {
            scope.InputValue = {};

            ngModel.$formatters.push(function (value) {
                scope.InputValue.Date = value;
                scope.InputValue.Time = value;
            });

            scope.calendarOpened = false;
            scope.calendarFormat = "yyyy-MM-dd";

            scope.openCalendar = function ($event) {
                $event.preventDefault();
                $event.stopPropagation();
                scope.calendarOpened = true;
            }

            scope.dateTimeChanged = function () {
                var dateTimeValue = new Date();
                if (scope.InputValue.Date !== undefined) {
                    dateTimeValue.setFullYear(scope.InputValue.Date.getFullYear());
                    dateTimeValue.setMonth(scope.InputValue.Date.getMonth());
                    dateTimeValue.setDate(scope.InputValue.Date.getDate());
                }
                if (scope.InputValue.Time !== undefined) {
                    dateTimeValue.setHours(scope.InputValue.Time.getHours());
                    dateTimeValue.setMinutes(scope.InputValue.Time.getMinutes());
                }
                dateTimeValue.setSeconds(0);
                ngModel.$setViewValue(dateTimeValue);
            }

            var outerElement = $("<span></span>");
            outerElement.attr("class", "input-group");

            var dateElement = $("<span></span>");
            dateElement.attr("class", "input-group");

            var dateTextInputElement = $("<input />");
            dateTextInputElement.attr("class", "form-control");
            dateTextInputElement.attr("data-ng-model", "InputValue.Date");
            dateTextInputElement.attr("data-datepicker-popup", "{{calendarFormat}}");
            dateTextInputElement.attr("data-is-open", "calendarOpened");
            dateTextInputElement.attr("data-ng-change", "dateTimeChanged()");
            //dateTextInputElement.attr("data-min-date", "'2015-01-01'");
            //dateTextInputElement.attr("data-max-date", "'2015-06-22'");

            dateElement.append(dateTextInputElement);

            var dateButtonContainerElement = $("<span></span>");
            dateButtonContainerElement.attr("class", "input-group-btn");

            var dateButtonInputElement = $("<button></button>");
            dateButtonInputElement.attr("class", "btn btn-default");
            dateButtonInputElement.attr("data-ng-click", "openCalendar($event)");
            
            var dateButtonImageElement = $("<i></i>");
            dateButtonImageElement.attr("class", "glyphicon glyphicon-calendar");

            dateButtonInputElement.append(dateButtonImageElement);

            dateButtonContainerElement.append(dateButtonInputElement);
            dateElement.append(dateButtonContainerElement);

            outerElement.append(dateElement);

            var timeElement = $("<timepicker></timepicker>");
            timeElement.attr("data-ng-model", "InputValue.Time");
            timeElement.attr("data-ng-change", "dateTimeChanged()");

            outerElement.append(timeElement);

            $compile(outerElement)(scope);
            return outerElement;
        }

        function generatePeoplePickerElement(element, ngModel, scope) {

            if (ngModel.$viewValue !== undefined) {
                scope.InputValue = ngModel.$viewValue;
            }
            else {
                scope.InputValue = null;
            }

            var elementId = "pp_" + scope.ngspColumn.Title + "_" + scope.ngspColumn.Id;
            var outerElement = $("<span></span>");
            outerElement.attr("id", elementId);

            $compile(outerElement)(scope);

            var schema = {
                SearchPrincipalSource: 15,
                ResolvePrincipalSource: 15,
                MaximumEntitySuggestions: 50,
                Width: "100%",
                OnUserResolvedClientScript: onUserResolve
            };

            schema.Required = scope.ngspColumn.IsRequired

            if (scope.ngspColumn.PeopleOnly) {
                schema.PrincipalAccountType = "User,DL";
            }
            else {
                schema.PrincipalAccountType = "User,DL,SecGroup,SPGroup";
            }

            if (scope.ngspColumn.Type === "person") {
                schema.AllowMultipleValues = false;
            }
            else if (scope.ngspColumn.Type === "multiPerson") {
                schema.AllowMultipleValues = true;
            }
            else {
                console.log(scope.ngspColumn.Type + "is not supported as a column type for people pickers.");
                return;
            }

            

            scope.$watch(element[0].childNodes.length, function () {
                var ppElement = $("#" + elementId);
                if (ppElement.length > 0) {
                    SPClientPeoplePicker_InitStandaloneControlWrapper(elementId, scope.InputValue, schema);
                    ngModel.$formatters.push(function (value) {
                        //ppElement.html("");
                        SPClientPeoplePicker_InitStandaloneControlWrapper(elementId, value, schema);
                    });
                }
            });

            function onUserResolve() {
                var peoplePickerDictKey = elementId + "_TopSpan";
                var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerDictKey];
                scope.InputValue = peoplePicker.GetAllUserInfo();
                //console.log(scope.InputValue);
                ngModel.$setViewValue(scope.InputValue);
            }

            return outerElement;
        }

        function generateDisplayElement(ngModel, scope) {
            switch (scope.ngspColumn.Type) {
                case "text":
                    return generateTextDisplayElement(ngModel, scope);
                    break;
                case "multiText":
                    return generateMultiTextDisplayElement(ngModel, scope);
                    break;
                case "choice":
                    return generateChoiceDisplayElement(ngModel, scope);
                    break;
                case "multiChoice":
                    return generateMultiChoiceDisplayElement(ngModel, scope);
                    break;
                case "lookup":
                    return generateLookupDisplayElement(ngModel, scope);
                    break;
                case "multiLookup":
                    return generateMultiLookupDisplayElement(ngModel, scope);
                    break;
                case "yesNo":
                    return generateYesNoDisplayElement(ngModel, scope);
                    break;
                case "person":
                case "multiPerson":
                    return generatePersonDisplayElement(ngModel, scope);
                    break;
                    break;
                case "metadata":
                    return generateMetadataDisplayElement(ngModel, scope);
                    break;
                case "multiMetadata":
                    return generateMultiMetadataDisplayElement(ngModel, scope);
                    break;
                case "link":
                    return generateLinkDisplayElement(ngModel, scope);
                    break;
                case "number":
                    return generateNumberDisplayElement(ngModel, scope);
                    break;
                case "currency":
                    return generateCurrencyDisplayElement(ngModel, scope);
                    break;
                case "dateTime":
                    return generateDateTimeDisplayElement(ngModel, scope);
                    break;
                default:
                    console.log(scope.ngspColumn.Type + " is currently unsupported as a column type.");
                    break;
            }

            function generateTextDisplayElement(ngModel, scope) {
                scope.DisplayValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = value;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }

            function generateMultiTextDisplayElement(ngModel, scope) {
                scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = value;
                });

                var outerElement = $("<span></span>");
                outerElement.attr("data-ng-bind-html", "DisplayValue");

                $compile(outerElement)(scope);

                return outerElement;
            }

            function generateChoiceDisplayElement(ngModel, scope) {
                scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = value;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }

            function generateMultiChoiceDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    var choices = [];
                    for (var key in value) {
                        if (value[key]) {
                            choices.push(key);
                        }
                    }
                    scope.DisplayValue = choices.join("; ");
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }
            
            function generateLookupDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = scope.ngspColumn.LookupItemsById[value].label;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }

            function generateMultiLookupDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    var lookupLabels = [];
                    for (var key in value) {
                        if (value[key]) {
                            var lookupLabel = scope.ngspColumn.LookupItemsById[key].label;
                            lookupLabels.push(lookupLabel);
                        }
                    }
                    scope.DisplayValue = lookupLabels.join("; ");
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }

            function generateYesNoDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    if (value) {
                        scope.DisplayValue = "Yes";
                    }
                    else {
                        scope.DisplayValue = "No";
                    }
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }

            function generatePersonDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    var people = [];
                    value.forEach(function (person) {
                        people.push(person.AutoFillDisplayText);
                    });
                    scope.DisplayValue = people.join("; ");
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }
            
            function generateMetadataDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = scope.ngspColumn.TermsById[value].label;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }
            
            function generateMultiMetadataDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    var metadataLabels = [];
                    for (var key in value) {
                        if (value[key]) {
                            var metadataLabel = scope.ngspColumn.TermsById[key].label;
                            metadataLabels.push(metadataLabel);
                        }
                    }
                    scope.DisplayValue = metadataLabels.join("; ");
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }
            
            function generateLinkDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = value;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }
            
            function generateNumberDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;
                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = value;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }
            
            function generateCurrencyDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = value;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }

            function generateDateTimeDisplayElement(ngModel, scope) {
                //scope.InputValue = ngModel.$viewValue;

                ngModel.$formatters.push(function (value) {
                    scope.DisplayValue = value;
                });

                var outerElement = $("<span></span>");
                outerElement.text("{{DisplayValue}}");

                $compile(outerElement)(scope);

                return outerElement;
            }
        }
    }

    function angularSpList($compile) {
        return {
            restrict: "A",
            require: "ngModel",
            scope: {
                ngspList: '=',
                ngspReadOnly: '='
            },
            link: link
        }

        function link(scope, element, attrs, ngModel) {
            scope.ngspList = {};
            scope.$watch('ngspList', function () {
                if (scope.ngspList.DisplayName !== undefined && scope.ngspList.Columns !== undefined) {
                    element.html("");
                    element.append(generateTitleElement(scope.ngspList));
                    element.append(generateInputElement(ngModel, scope));
                    element.append(generateButtonsElement(scope));
                    //element.append(generateButtonsElement(ngModel, scope));
                }
            });
        }

        function generateTitleElement(list) {
            var titleElement = $("<label></label>");
            titleElement.text(list.DisplayName);
            return titleElement;
        }

        function generateInputElement(ngModel, scope) {
            //switch (scope.ngspList.Type) {
            //    case "Calendar":
            //        console.log("Calendar list type is currently not supported.");
            //        return;
            //        break;
            //    default:
            return generateGridElement(ngModel, scope);
            //return generateTableElement(ngModel, scope);
            //        break;
            //}
        }

        function generateTableElement(ngModel, scope) {
            scope.InputValues = [
                {
                    "One": "One",
                    "Two": "2"
                }
            ];

            var outerElement = $("<span></span>");
            var tableElement = $("<table></table>");
            var headerElement = $("<th></th>");
            var rowElement = $("<tr></tr>");
            
            var columns = scope.ngspList.Columns;
            for (var columnName in columns) {
                if (columns.hasOwnProperty(columnname)) {
                    var column = columns[columnName];
                    var headerColumnElement = $("<td></td>");
                    if (column.DisplayName !== undefined) {
                        headerColumnElement.text("{{ngspList.Columns[" + columnName + "].DisplayName}}");
                    }
                    else {
                        headerColumnElement.text(columnName);
                    }
                    headerElement.append(headerColumnElement);
                    var rowColumnElement = $("<td></td>");
                    rowColumnElement.text("{{InputValues[" + columnName + "]}}");
                    rowElement.append(rowColumnElement);
                }
            }

            outerElement.append()

            $compile(outerElement)(scope);
            return outerElement;
        }

        function generateGridElement(ngModel, scope) {
            scope.InputValues = [
                {
                    "One": "One",
                    "Two": "2"
                }
            ];

            scope.uiGridOptions = {
                data: scope.InputValues,
                enableRowSelection: true,
                enableRowHeaderSelection: false,
                multiSelect: false,
                enableHorizontalScrollbar: 1, 
                enableVerticalScrollbar: 1,
                onRegisterApi: registerGridApi,
            }

            function registerGridApi(gridApi) {
                scope.gridApi = gridApi;
                console.log(scope.gridApi);
                scope.gridApi.selection.on.rowSelectionChanged(scope, function () {
                    scope.selectedRow = scope.gridApi.selection.getSelectedRows()[0];
                });
            }

            var outerElement = $("<span></span>");
            var gridElement = $("<div></div>");
            gridElement.attr("data-ui-grid", "uiGridOptions");
            gridElement.attr("data-ui-grid-selection", "true");
            gridElement.attr("class", "grid");
            outerElement.append(gridElement);

            $compile(outerElement)(scope);
            return outerElement;
        }

        function generateButtonsElement(scope) {
            scope.createRow = function () {
                scope.InputValues.push({
                    "One": "Another",
                    "Two": "Second"
                });
            }

            scope.updateRow = function () {
                if (scope.selectedRow === undefined) {
                    alert("Nothing selected");
                }
                else {

                }
            }

            scope.deleteRow = function () {
                if (scope.selectedRow === undefined) {
                    alert("Nothing selected");
                }
                else {

                }
            }

            scope.rowNotSelected = function () {
                if (scope.selectedRow === undefined) {
                    return true;
                }
                else {
                    return false;
                }
            }

            var outerElement = $("<span></span>");

            var createElement = $("<input></input>");
            createElement.attr("type", "button");
            createElement.attr("value", "Add");
            createElement.attr("data-ng-click", "createRow()");
            outerElement.append(createElement);

            var updateElement = $("<input></input>");
            updateElement.attr("type", "button");
            updateElement.attr("value", "Edit");
            updateElement.attr("data-ng-disabled", "rowNotSelected()");
            updateElement.attr("data-ng-click", "updateRow()");
            outerElement.append(updateElement);

            var deleteElement = $("<input></input>");
            deleteElement.attr("type", "button");
            deleteElement.attr("value", "Delete");
            deleteElement.attr("data-ng-disabled", "rowNotSelected()");
            deleteElement.attr("data-ng-click", "deleteRow()");
            outerElement.append(deleteElement);

            $compile(outerElement)(scope);
            return outerElement;
        }
    }

})(angular, jQuery)