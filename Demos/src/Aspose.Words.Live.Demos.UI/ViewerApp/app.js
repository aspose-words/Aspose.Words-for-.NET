var app = angular.module('AsposeViewerApp', [
	'ngSanitize'
]);

app.config(['$httpProvider',
	function ($httpProvider) {
		$httpProvider.defaults.cache = false;
		if (!$httpProvider.defaults.headers.get) 
			$httpProvider.defaults.headers.get = {};
		$httpProvider.defaults.headers.get['Cache-Control'] = 'no-cache';
		$httpProvider.defaults.headers.get['Pragma'] = 'no-cache';
	}]);

app.run(function ($rootScope, $templateCache) {
	$rootScope.$on('$viewContentLoaded', function () {
		$templateCache.removeAll();
	});
});
