var app = app || angular.module("angularSP", ['ui.bootstrap', 'ui.grid', 'ui.grid.selection']);

(function (ng, $) {
    app.controller('listItemModalController', ["$scope", "$modalInstance", "item", "columns", listItemModalCtrl]);

    function listItemModalCtrl($scope, $modalInstance, item, columns) {
        $scope.item = item;
        $scope.columns = columns;

        $scope.submit = function (newItem) {
            $modalInstance.close(newItem);
        }

        $scope.cancel = function () {
            $modalInstance.dismiss("Canceled.");
        }
    }
})(angular, jQuery)