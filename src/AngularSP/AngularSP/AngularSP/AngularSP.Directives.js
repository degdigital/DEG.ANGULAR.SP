'use strict';

(function (ng, $) {
    ng.module("angularSP")
        .directive("contenteditable", function () {
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

    ng.module("angularSP")
        .directive("ngspPeoplepicker", function () {
            return {
                restrict: "A",
                require: "?ngModel",
                link: function (scope, element, attrs, ngModel) {
                    if (!ngModel) {
                        return;
                    }
                    var elementId = attrs.id;

                    var schema = {
                        PrincipalAccountType: "User,DL,SecGroup,SPGroup",
                        SearchPrincipalSource: 15,
                        ResolvePrincipalSource: 15,
                        MaximumEntitySuggestions: 50,
                        Width: "100%",
                    };

                    if (scope.ngspColumn.IsRequired) {
                        schema.Required = true;
                    }
                    else {
                        schema.Required = false;
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

                    SPClientPeoplePicker_InitStandaloneControlWrapper(elementId, null, schema);
                }
            };
        });

    ng.module("angularSP")
        .directive("ngspColumn", angularSpField);
    
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
                    element.append(generateInputElement(scope));
                    if (scope.ngspColumn.InputType === "nicEdit") {
                        
                    }
                }
            });
        }

        function generateTitleElement(column) {
            var titleElement = $("<label></label>");
            titleElement.text(column.Title);
            return titleElement;
        }

        function generateInputElement(scope) {
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
                case "peoplePicker":
                    return generatePeoplePickerElement(scope);
                    break;
                default:
                    console.log(scope.ngspColumn.InputType + " is currently unsupported as an input type.");
                    break;
            }
        }

        function generateTextElement(scope) {
            var inputElement = $("<input />");
            inputElement.attr("data-ng-model", "ngspColumn.Value");
            $compile(inputElement)(scope);
            return inputElement;
        }

        function generateTextBoxElement(scope) {
            var inputElement = $("<textarea></textarea>");
            inputElement.attr("data-ng-model", "ngspColumn.Value");
            $compile(inputElement)(scope);
            return inputElement;
        }

        function generateNicEditElement(scope) {
            var outerElement = $("<span></span>");
            var panelElement = $("<div></div>");
            //panelElement.attr("id", scope.ngspColumn.Title + "_panel");
            outerElement.append(panelElement);
            var contentElement = $("<div></div>");
            //panelElement.attr("id", scope.ngspColumn.Title + "_content");
            contentElement.attr("contenteditable", "true");
            contentElement.attr("data-ng-model", "ngspColumn.Value");
            outerElement.append(contentElement);

            var nicEditorConstrutorOption = { iconsPath: '/AngularSP/lib/images/nicEditorIcons.gif' };
            var nicEditorInstance = new nicEditor(nicEditorConstrutorOption);
            nicEditorInstance.setPanel(panelElement[0]);
            nicEditorInstance.addInstance(contentElement[0]);

            $compile(outerElement)(scope);

            return outerElement;
        }

        function generateDropDownElement(scope) {
            var inputElement = $("<select></select>");
            inputElement.attr("data-ng-model", "ngspColumn.Value");
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
                    inputElement.attr("data-ng-model", "ngspColumn.ValueSelected[term.id]");
                    labelElement.text("{{term.label}}");
                }
                else if (scope.ngspColumn.Type === "multiLookup") {
                    repeatElement.attr("data-ng-repeat", "lookupItem in ngspColumn.LookupItems");
                    inputElement.attr("data-ng-model", "ngspColumn.ValueSelected[lookupItem.id]");
                    labelElement.text("{{lookupItem.label}}");
                }
                else if (scope.ngspColumn.Type === "multiChoice") {
                    repeatElement.attr("data-ng-repeat", "choice in ngspColumn.Choices");
                    inputElement.attr("data-ng-model", "ngspColumn.ValueSelected[choice]");
                    labelElement.text("{{choice}}");
                }
                repeatElement.append(inputElement);
                repeatElement.append(labelElement);
            }
            else if (scope.ngspColumn.Type === "yesNo") {
                var inputElement = $("<input />");
                inputElement.attr("type", "checkbox");
                inputElement.attr("data-ng-model", "ngspColumn.ValueSelected");
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

        function generatePeoplePickerElement(scope) {
            var elementId = "pp_" + scope.ngspColumn.Title;
            var outerElement = $("<span></span>");
            outerElement.attr("id", elementId);
            outerElement.attr("data-ngsp-peoplepicker", "true");
            outerElement.attr("data-ng-model", "ngspColumn.ValueSelected");

            $compile(outerElement)(scope);

            return outerElement;
        }
    }
})(angular, jQuery)