'use strict';

(function (ng) {
    ng.module('testForm', ['angularSP'])
        .controller('testFormController', ['$scope', '$http', '$q', 'spListFactory', testFormCont])

    function testFormCont($scope, $http, $q, spListFactory) {
        console.log("Initiating...");

        var listDefPromise = $http.get('ListDefinitions.js');
        var columnDefPromise = $http.get('ColumnDefinitions.js');

        $q.all([columnDefPromise, listDefPromise]).then(function (data) {
            var columnDefs = data[0].data;
            var listDefs = data[1].data;

            console.log(columnDefs);
            console.log(listDefs);
            
            spListFactory.initFactory(columnDefs, listDefs);
        }, function (reason) {
            console.log(reason);
        });
    }
})(angular);