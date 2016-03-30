var app = app || angular.module("angularSP", ['ui.bootstrap', 'ui.grid', 'ui.grid.selection']);

(function (ng, $) {

    app.factory("spListFactory", ['$rootScope', '$q', '$log', spListFactory]);

})(angular, jQuery)