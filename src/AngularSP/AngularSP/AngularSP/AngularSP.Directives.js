'use strict';

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
            scope: {
                ngspColumn: '=',
            },
            link: link
        }

        function link(scope, element, attrs) {
            scope.ngspColumn = {};
            scope.$watch('ngspColumn', function () {
                if (scope.ngspColumn.Title !== undefined && scope.ngspColumn.InputType !== undefined) {
                    element.html("");
                    element.append(generateTitleElement(scope.ngspColumn));
                    element.append(generateInputElement(element, scope));
                }
            });
        }

        function generateTitleElement(column) {
            var titleElement = $("<label></label>");
            titleElement.text(column.Title);
            return titleElement;
        }

        function generateInputElement(element, scope) {
            switch (scope.ngspColumn.InputType) {
                case "text":
                    return generateTextElement(scope);
                    break;
                case "textBox":
                    return generateTextBoxElement(scope);
                    break;
                case "nicEdit":
                    return generateNicEditElement(scope);
                    break;
                case "dropDown":
                    return generateDropDownElement(scope);
                    break;
                case "checkbox":
                    return generateCheckboxElement(scope);
                    break;
                case "dateTime":
                    return generateDateTimeElement(scope);
                    break;
                case "peoplePicker":
                    return generatePeoplePickerElement(element, scope);
                    break;
                default:
                    console.log(scope.ngspColumn.InputType + " is currently unsupported as an input type.");
                    break;
            }
        }

        function generateTextElement(scope) {
            var inputElement = $("<input />");
            inputElement.attr("data-ng-model", "ngspColumn.InputValue");
            $compile(inputElement)(scope);
            return inputElement;
        }

        function generateTextBoxElement(scope) {
            var inputElement = $("<textarea></textarea>");
            inputElement.attr("data-ng-model", "ngspColumn.InputValue");
            $compile(inputElement)(scope);
            return inputElement;
        }

        function generateNicEditElement(scope) {
            var outerElement = $("<span></span>");
            var panelElement = $("<div></div>");
            outerElement.append(panelElement);
            var contentElement = $("<div></div>");
            contentElement.attr("contenteditable", "true");
            contentElement.attr("data-ng-model", "ngspColumn.InputValue");
            outerElement.append(contentElement);


            //JORDAN: parameterize this.
            var nicEditorConstrutorOption = { iconsPath: '/AngularSP/lib/images/nicEditorIcons.gif' };
            var nicEditorInstance = new nicEditor(nicEditorConstrutorOption);
            nicEditorInstance.setPanel(panelElement[0]);
            nicEditorInstance.addInstance(contentElement[0]);

            $compile(outerElement)(scope);

            return outerElement;
        }

        function generateDropDownElement(scope) {
            var inputElement = $("<select></select>");
            inputElement.attr("data-ng-model", "ngspColumn.InputValue");
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

        function generateCheckboxElement(scope) {
            var outerElement = $("<span></span>");
            if (scope.ngspColumn.Type === "multiMetadata" || scope.ngspColumn.Type === "multiLookup" || scope.ngspColumn.Type === "multiChoice") {
                var repeatElement = $("<span></span>");
                var inputElement = $("<input />");
                var labelElement = $("<label></label>");
                inputElement.attr("type", "checkbox");
                if (scope.ngspColumn.Type === "multiMetadata")
                {
                    repeatElement.attr("data-ng-repeat", "term in ngspColumn.Terms");
                    inputElement.attr("data-ng-model", "ngspColumn.InputValue[term.id]");
                    labelElement.text("{{term.label}}");
                }
                else if (scope.ngspColumn.Type === "multiLookup") {
                    repeatElement.attr("data-ng-repeat", "lookupItem in ngspColumn.LookupItems");
                    inputElement.attr("data-ng-model", "ngspColumn.InputValue[lookupItem.id]");
                    labelElement.text("{{lookupItem.label}}");
                }
                else if (scope.ngspColumn.Type === "multiChoice") {
                    repeatElement.attr("data-ng-repeat", "choice in ngspColumn.Choices");
                    inputElement.attr("data-ng-model", "ngspColumn.InputValue[choice]");
                    labelElement.text("{{choice}}");
                }
                repeatElement.append(inputElement);
                repeatElement.append(labelElement);
            }
            else if (scope.ngspColumn.Type === "yesNo") {
                var inputElement = $("<input />");
                inputElement.attr("type", "checkbox");
                inputElement.attr("data-ng-model", "ngspColumn.InputValue");
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

        function generateDateTimeElement(scope) {
            scope.calendarOpened = false;
            scope.calendarFormat = "yyyy-MM-dd";

            scope.openCalendar = function ($event) {
                $event.preventDefault();
                $event.stopPropagation();
                scope.calendarOpened = true;
            }

            var outerElement = $("<span></span>");
            outerElement.attr("class", "input-group");

            var dateElement = $("<span></span>");
            dateElement.attr("class", "input-group");

            var dateTextInputElement = $("<input />");
            dateTextInputElement.attr("class", "form-control");
            dateTextInputElement.attr("data-ng-model", "ngspColumn.InputValue.Date");
            dateTextInputElement.attr("data-datepicker-popup", "{{calendarFormat}}");
            dateTextInputElement.attr("data-is-open", "calendarOpened");
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
            timeElement.attr("data-ng-model", "ngspColumn.InputValue.Time");

            outerElement.append(timeElement);

            $compile(outerElement)(scope);
            return outerElement;
        }

        function generatePeoplePickerElement(element, scope) {
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
                if ($("#" + elementId).length > 0) {
                    SPClientPeoplePicker_InitStandaloneControlWrapper(elementId, null, schema);
                }
            });

            function onUserResolve() {
                var peoplePickerDictKey = elementId + "_TopSpan";
                var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerDictKey];
                scope.ngspColumn.InputValue = peoplePicker.GetAllUserInfo();
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