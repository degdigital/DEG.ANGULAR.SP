
var app = app || angular.module("angularSP", ['ui.bootstrap']);

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
    
    function angularSpField ($compile) {
        return {
            restrict: "A",
            require: "ngModel",
            scope: {
                ngspColumn: '=',
                ngModel: '='
            },
            link: link
        }

        function link(scope, element, attrs, ngModel) {
            scope.ngspColumn = {};
            scope.$watch('ngspColumn', function () {
                if (scope.ngspColumn.Title !== undefined && scope.ngspColumn.InputType !== undefined) {
                    element.html("");
                    element.append(generateTitleElement(scope.ngspColumn));
                    element.append(generateInputElement(element, ngModel, scope));
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

            var outerElement = $("<span></span>");

            if (scope.ngspColumn.Type === "yesNo") {
                console.log("Radio input is not implemented yet for yesNo.");
                return;
            }
            else if (scope.ngspColumn.Type === "choice") {
                console.log("Radio input is not implemented yet for choice.");
                return;
            }
            else if (scope.ngspColumn.Type === "lookup") {
                console.log("Radio input is not implemented yet for lookup.");
                return;
            }
            else if (scope.ngspColumn.Type === "metadata") {
                console.log("Radio input is not implemented yet for metadata.");
                return;
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
    }

    app.directive("ngspList", angularSpList);

    function angularSpList($compile) {
        return {
            restrict: "A",
            scope: {
                ngspList: '=',
            },
            link: link
        }

        function link(scope, element, attrs) {
            scope.ngspList = {};
            scope.$watch('ngspList', function () {
                if (scope.ngspList.DisplayName !== undefined && scope.ngspList.Columns !== undefined) {
                    element.html("");
                    element.append(generateTitleElement(scope.ngspList));
                    element.append(generateGridElement());
                    element.append(generateButtonsElement(element, scope));
                }
            });
        }
    }

})(angular, jQuery)