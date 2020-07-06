angular.module('AsposeEditorApp').controller('EditorController', function ($scope, $http) {

	var initialized = false;
	$scope.trum = null; // jquery object
	$scope.t = null; // Trumbowyg component
	$scope.loading = false;

    window.history.pushState(null, null, '/' + productName + '/editor');

	// 'back' button special handler
	/*var onpopstatePrev = window.onpopstate;
	window.onpopstate = function(event) {
		event.preventDefault();
		event.stopPropagation();
		window.onpopstate = onpopstatePrev;
		window.location.href = '/words/editor';
	};*/

	$scope.Init = function () {
		if (initialized === false) {
			initialized = true;
			$scope.trum = $('#editor');
			$scope.trum.trumbowyg({
				autogrow: true,
				fixedBtnPane: true,
				minimalLinks: true,
				urlProtocol: true,
				prefix: "editoraspose-",
				svgPath: "/editor/Trumbowyg/icons.svg",
				btnsDef: {
					image: {
						dropdown: ['insertImage', 'base64'],
						ico: 'insertImage'
					}
				},
				btns: [
					['historyUndo', 'historyRedo'],
					['fontfamily'],
					['fontsize'],
					['strong', 'em', 'underline'],
					['foreColor', 'backColor'],
					['superscript', 'subscript'],
					['link'],
					['image'],
					['justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'],
					['unorderedList', 'orderedList'],
					['horizontalRule'],
					['fullscreen']
				],
				plugins: {
					fontfamily: {
						fontList: [
							{ name: 'Arial', family: 'Arial, Helvetica, sans-serif' },
							{ name: 'Arial Black', family: '\'Arial Black\', Gadget, sans-serif' },
							{ name: 'Comic Sans', family: '\'Comic Sans MS\', Textile, cursive, sans-serif' },
							{ name: 'Courier New', family: '\'Courier New\', Courier, monospace' },
							{ name: 'Georgia', family: 'Georgia, serif' },
							{ name: 'Impact', family: 'Impact, Charcoal, sans-serif' },
							{ name: 'Lucida Console', family: '\'Lucida Console\', Monaco, monospace' },
							{ name: 'Lucida Sans', family: '\'Lucida Sans Uncide\', \'Lucida Grande\', sans-serif' },
							{ name: 'Palatino', family: '\'Palatino Linotype\', \'Book Antiqua\', Palatino, serif' },
							{ name: 'Tahoma', family: 'Tahoma, Geneva, sans-serif' },
							{ name: 'Times New Roman', family: '\'Times New Roman\', Times, serif' },
							{ name: 'Trebuchet', family: '\'Trebuchet MS\', Helvetica, sans-serif' },
							{ name: 'Verdana', family: 'Verdana, Geneva, sans-serif' }
						]
					},
					fontsize: {
						sizeList: [
							'9',
							'14',
							'18',
							'22',
							'30'
						],
						allowCustomSize: true
					}
				}
			});
			$scope.trum.trumbowyg('html', '');
			$scope.t = $scope.trum.data('trumbowyg'); // Component
			$scope.GetEditorHtml();
		}
	};

	$scope.GetEditorHtml = function () {
		$http({

			method: 'POST',
			url: asposeEditorAPP + 'GetHTML',
			params: {				
				'folderName': folderName,
				'fileName': fileName
			},
			responseType: "application/json"			
		}).then(function (response) {
			$scope.trum.trumbowyg('html', response.data);
		}, function (error) {
			$scope.ShowError(error.data);
			$scope.trum.trumbowyg('html', '');
		}).finally(function () {
			$('#page-loading').fadeOut(600);
			$('#htmlloader').hide();
		});
	};

	$scope.Download = function (outputType) {
		$('#page-loading').show();
		$('#loader').show();
		$http({
			method: 'POST',
			url: asposeEditorAPP + 'UpdateContents',
			data: {				
				'htmldata': $scope.trum.trumbowyg('html'),
				'folderName': folderName,
				'fileName': fileName,
				'outputType': outputType
			},
			responseType: "application/json"
		}).then(function (response) {
			var data = response.data;
			if (data.StatusCode === 200)
				window.location = fileDownloadLink + "?fileName=" + data.FileName + "&folderName=" + data.FolderName;
			else
				$scope.ShowError(data.Status);
		}, function (error) {
			$scope.ShowError(error.data);
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