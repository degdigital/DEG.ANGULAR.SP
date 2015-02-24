'use strict';

(function (ng) {
    ng.module('testForm', ['angularSP'])
        .controller('testFormController', ['$scope', '$http', '$q', 'spListFactory', testFormCont])

    function testFormCont($scope, $http, $q, spListFactory) {
        console.log("Initiating...");

        var columnDefPromise = $http.get('ColumnDefinitions.js');
        var cTypeDefPromise = $http.get('ContentTypeDefinitions.js');
        var listDefPromise = $http.get('ListDefinitions.js');

        $q.all([columnDefPromise, cTypeDefPromise, listDefPromise]).then(function (data) {
            var columnDefs = data[0].data;
            var cTypeDefs = data[1].data;
            var listDefs = data[2].data;

            console.log(columnDefs);
            console.log(cTypeDefs);
            console.log(listDefs);
            
            spListFactory.initFactory(columnDefs, cTypeDefs, listDefs);
        }, function (reason) {
            console.log(reason);
        });
    }
})(angular);