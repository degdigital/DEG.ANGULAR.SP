'use strict';

(function (ng) {
    ng.module('testForm', ['angularSP'])
        .controller('testFormController', ['$rootScope', '$scope', '$http', '$window', '$q', 'spListFactory', testFormCont]);

    function testFormCont($rootScope, $scope, $http, $window, $q, spListFactory) {
        $scope.showForm = false;

        $scope.AngularSPTestList = {};

        var queryStrings = getQueryStrings();

        console.log("Initiating...");

        var fn = function () {

        }

        console.log(typeof (fn));

        var columnDefPromise = $http.get('ColumnDefinitions.js');
        var cTypeDefPromise = $http.get('ContentTypeDefinitions.js');
        var listDefPromise = $http.get('ListDefinitions.js');

        $q.all([columnDefPromise, cTypeDefPromise, listDefPromise]).then(function (data) {
            var columnDefs = data[0].data;
            var cTypeDefs = data[1].data;
            var listDefs = data[2].data;

            //console.log(columnDefs);
            //console.log(cTypeDefs);
            //console.log(listDefs);
            
            //console.log(spListFactory.getServerRelativeUrl());

            var initPromise = spListFactory.initFactory(columnDefs, cTypeDefs, listDefs);

            initPromise.then(initSuccess, initFailure);

            function initSuccess(listsInfo) {
                $scope.listsInfo = listsInfo;
                console.log($scope.listsInfo);
                if (queryStrings.ID !== undefined) {
                    var listItemPromise = spListFactory.getListItem("AngularSPTestList", queryStrings.ID);
                    listItemPromise.then(function (itemDetails) {
                        $scope.AngularSPTestList = itemDetails;
                        //$scope.$apply();
                        console.log($scope.AngularSPTestList);
                        $scope.showForm = true;
                    }, function (reason) {
                        console.log(reason);
                    });
                }
                else {
                    $scope.showForm = true;
                }
            }

            function initFailure(reason) {
                console.log(reason);
            }

        }, function (reason) {
            console.log(reason);
        });

        $scope.Submit = function (model) {
            console.log(model);
            var createItemPromise = spListFactory.createListItem("AngularSPTestList", model);

            createItemPromise.then(function (itemId) {
                console.log(itemId);
            }, function (reason) {
                console.log(reason);
            });
        };

        function getQueryStrings() {
            var queryString = $window.location.search;

            var pairs = queryString.substring(1).split("&");

            var queryStrings = {};

            for (var i = 0; i < pairs.length; i++) {
                if (pairs[i] !== "") {
                    var pair = pairs[i].split("=");
                    if (pair.length === 2) {
                        queryStrings[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1]);
                    }
                }
            }
            
            return queryStrings;
        }
    }
})(angular);