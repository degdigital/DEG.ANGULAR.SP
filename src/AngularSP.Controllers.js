var app = app || angular.module("angularSP", ['ui.bootstrap', 'ui.grid', 'ui.grid.selection']);

(function (ng, $) {
    app.controller('listItemModalController', listItemModalCtrl);

    function listItemModalCtrl($scope, $modalInstance, items) {
        $scope.items = items;
    }
})(angular, jQuery)