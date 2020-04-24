

angular.module('AsposeViewerApp').controller('ViewerController', function ($scope, $http, $q) {

	var initialized = false;
	$scope.loading = false;

	$scope.Init = function () {
		if (initialized === false) {
			initialized = true;
			$.support.cors = true;
			$('#viewer').viewer({
				applicationPath: asposeViewerAPI,
				documentGuid: fileName,
				folderName: folderName,
				defaultDocument: fileName,
				htmlMode: true,
				preloadPageCount: 3,
				zoom: true,
				pageSelector: true,
				search: true,
				thumbnails: true,
				rotate: true,
				download: false,
				upload: false,
				print: true,
				browse: false,
				rewrite: true,
				saveRotateState: false,
				enableRightClick: false
			});
			$('#gd-pages').on('click', 'a', function (el) {
				var h = $(this).attr('href');
				if (h.startsWith('#_page')) {
					var pageNumber = parseInt(h.split('_')[1].substring(4));
					appendHtmlContent(pageNumber, fileName, function() {
						window.location.href = h;
						setNavigationPageValues(pageNumber, totalPageNumber);
						if (pageNumber > 1) { // load previous page
							appendHtmlContent(pageNumber - 1, fileName, function() { // load next page
								if (pageNumber < totalPageNumber) {
									appendHtmlContent(pageNumber + 1, fileName, function() {});
								}
							});
						}
					});
				} else
					window.open(h);
				return false;
			});
			
			// var beforeQueryString = window.location.href.split("?")[0];
			// window.history.pushState({}, document.title, beforeQueryString);
		}
	};
	
	$scope.Download = function (outputType) {
		$('#page-loading').show();
		$('#loader').show();
		$http({
			type: 'POST',	
			url: asposeViewerAPI + 'Download' + "?fileName=" + fileName + "&folderName=" + folderName + "&outputType=" + outputType,			
			//data: {
			//	'folderName': folderName,
			//	'fileName': fileName,
			//	'outputType': outputType
			//},			
			responseType: "application/json"
		}).then(function (response) {
			var data = response.data;
			if (data.StatusCode === 200)
				window.location = fileDownloadLink + "?fileName=" + data.FileName + "&folderName=" + data.FolderName;
			else
				$scope.ShowError(data.Status);
		}, function (error) {
			$scope.ShowError(error.message);
		}).finally(function () {
			$('#page-loading').fadeOut(600);
		});
	};

	$scope.ShowError = function (message) {
		$('#alert > p')[0].innerText = message;
		$('#alert').show();
	};

	$scope.Init();
});