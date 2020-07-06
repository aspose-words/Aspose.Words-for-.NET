const VIEWABLE_EXTENSIONS = [
	'DOC', 'DOCX', 'DOCM', 'DOT', 'DOTX', 'DOTM',
	'RTF', 'HTML', 'MHTML',
	'MOBI', 'ODT', 'OTT', 'TXT'
];


// filedrop component
var fileDrop = {};
var fileDrop2 = {};

$.extend($.expr[':'], {
	isEmpty: function (e) {
		return e.value === '';
	}
});

// Restricts input for the set of matched elements to the given inputFilter function.
(function ($) {
	$.fn.inputFilter = function (inputFilter) {
		return this.on("input keydown keyup mousedown mouseup select contextmenu drop", function () {
			if (inputFilter(this.value)) {
				this.oldValue = this.value;
				this.oldSelectionStart = this.selectionStart;
				this.oldSelectionEnd = this.selectionEnd;
			} else if (this.hasOwnProperty("oldValue")) {
				this.value = this.oldValue;
				this.setSelectionRange(this.oldSelectionStart, this.oldSelectionEnd);
			} else {
				this.value = "";
			}
		});
	};
}(jQuery));

function showLoader() {
	$('.progress > .progress-bar').html('15%');
	$('.progress > .progress-bar').css('width', '15%');
	$('#loader').removeClass("hidden");
	hideAlert();
}

function hideLoader() {
	$('#loader').addClass("hidden");
}
function generateViewerLink(data) {

	var response = null;
	var statusCode = null;
	var fileName = null;
	var folderName = null;
	var fileProcessingErrorCode = null;
	if (data.FileName !== undefined) {

		statusCode = data.StatusCode;
		fileName = data.FileName;
		folderName = data.FolderName;
		fileProcessingErrorCode = data.FileProcessingErrorCode;
	}
	else {
		response = data.split('|');
		statusCode = response[0];
		fileName = response[1];
		folderName = response[2];
		fileProcessingErrorCode = response[3];
	}
	

	//var id = data.FolderName !== undefined ? data.FolderName : data.id;
	return encodeURI(o.ViewerPath +
		'FileName=' +
		fileName +
		'&FolderName=' +
		folderName +
		'&CallbackURL=' +
		o.AppURL);
}
function generateEditorLink(data) {
	var response = data.split('|');
	return encodeURI(o.EditorPath +
		'FileName=' +
		response[1] +
		'&FolderName=' +
		response[2] +
		'&CallbackURL=' +
		o.AppURL);
}
function sendPageView(url) {
	if ('ga' in window)
		try {
			var tracker = ga.getAll()[0];
			if (tracker !== undefined) {
				tracker.send('pageview', url);
			}
		} catch (e) {
			/**/
		}
}
function workSuccess(data, textStatus, xhr) {
	hideLoader();
	var response = null;
	var statusCode = null;
	var fileName = null;
	var folderName = null;
	var fileProcessingErrorCode = null;
	if (data.FileName !== undefined) {

		statusCode = data.StatusCode;
		fileName = data.FileName;
		folderName = data.FolderName;
		fileProcessingErrorCode = data.FileProcessingErrorCode;
	}
	else {
		 response = data.split('|');
		 statusCode = response[0];
		 fileName = response[1];
		 folderName = response[2];
		 fileProcessingErrorCode = response[3];
	}
	
	if (fileName != null) {
		
		if (statusCode == '200') {

			if (fileProcessingErrorCode !== undefined && fileProcessingErrorCode !="0") {
				showAlert(o.FileProcessingErrorCodes[fileProcessingErrorCode]);
				return;
			}

			$('#WorkPlaceHolder').addClass('hidden');
			$('#DownloadPlaceHolder').removeClass('hidden');
			//if (o.ReturnFromViewer === undefined) {
			//	const pos = o.AppDownloadURL.indexOf('?');
			//	const url = pos === -1 ? o.AppDownloadURL : o.AppDownloadURL.substring(0, pos);
			//	sendPageView(url);
			//}
			var url = encodeURI(o.UIBasePath + `common/download?fileName=${fileName}&folderName=${folderName}`);

			$('#DownloadButton').attr('href', url);
			o.DownloadUrl = url;

			if (o.ShowViewerButton) {
				let viewerlink = $('#ViewerLink');
				let dotPos = fileName.lastIndexOf('.');
				let ext = dotPos >= 0 ? fileName.substring(dotPos + 1).toUpperCase() : null;
				if (ext !== null && viewerlink.length && VIEWABLE_EXTENSIONS.indexOf(ext) !== -1) {
					viewerlink.on('click', function (evt) {
						evt.preventDefault();
						evt.stopPropagation();
						openIframe(generateViewerLink(data), '/words/viewer', '/words/view');
					});
				}
				else {
					viewerlink.hide();
					$(viewerlink[0].parentNode.previousElementSibling).hide(); // div.clearfix	
				}
			}
		}
		else {
			showAlert(statusCode);
		}
	}
}



function hideAlert() {
	$('#alertMessage').addClass("hidden");
	$('#alertMessage').text("");
	$('#alertSuccess').addClass("hidden");
	$('#alertSuccess').text("");
}

function showAlert(msg) {
	hideLoader();
	$('#alertMessage').html(msg);
	$('#alertMessage').removeClass("hidden");
	$('#alertMessage').fadeOut(100).fadeIn(100).fadeOut(100).fadeIn(100);
}

function showMessage(msg) {
	hideLoader();
	$('#alertSuccess').text(msg);
	$('#alertSuccess').removeClass("hidden");
}

(function ($) {
	$.QueryString = (function (paramsArray) {
		let params = {};

		for (let i = 0; i < paramsArray.length; ++i) {
			let param = paramsArray[i]
				.split('=', 2);

			if (param.length !== 2)
				continue;

			params[param[0]] = decodeURIComponent(param[1].replace(/\+/g, " "));
		}

		return params;
	})(window.location.search.substr(1).split('&'))
})(jQuery);

function progress(evt) {
	if (evt.lengthComputable) {
		var max = evt.total;
		var current = evt.loaded;

		var percentage = Math.round((current * 100) / max);
		percentage = (percentage < 15 ? 15 : percentage) + '%';

		$('.progress > .progress-bar').html(percentage);
		$('.progress > .progress-bar').css('width', percentage);
	}
}

function removeAllFileBlocks() {
	fileDrop.droppedFiles.forEach(function (item) {
		$('#fileupload-' + item.id).remove();
	});
	fileDrop.droppedFiles = [];
	hideLoader();
}


function openIframe(url, fakeUrl, pageViewUrl) {
	// push fake state to prevent from going back
	window.history.pushState(null, null, fakeUrl);

	// remove body scrollbar
	$('body').css('overflow-y', 'hidden');

	// create iframe and add it into body
	var div = $('<div id="iframe-wrap"></div>');
	$('<iframe>', {
		src: url,
		id: 'iframe-document',
		frameborder: 0,
		scrolling: 'yes'
	}).appendTo(div);
	div.appendTo('body');
	sendPageView(pageViewUrl);
}

function closeIframe() {
	removeAllFileBlocks();
	$('div#iframe-wrap').remove();
	$('body').css('overflow-y', 'auto');
}
function request(url, data) {
	showLoader();
	$.ajax({
		type: 'POST',
		url: url,
		data: data,
		cache: false,
		contentType: false,
		processData: false,
		success: workSuccess,		
		xhr: function () {
			var myXhr = $.ajaxSettings.xhr();
			if (myXhr.upload)
				myXhr.upload.addEventListener('progress', progress, false);
			return myXhr;
		},
		error: function (err) {
			if (err.data !== undefined && err.data.Status !== undefined)
				showAlert(err.data.Status);
			else
				showAlert("Error " + err.status + ": " + err.statusText);
		}
	});
}
function requestMerger() {
	let data = fileDrop.prepareFormData(2, o.MaximumUploadFiles);
	if (data === null)
		return;


	let url = o.UIBasePath + 'Merger/Merger?outputType=' + $('#saveAs').val();
	request(url, data);
}
function requestParser() {
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;

	let url = o.UIBasePath + 'Parser/Parser';
	request(url, data);
}
function requestAnnotation() {
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;
	let url = o.UIBasePath + 'Annotation/Remove';
	request(url, data);
}	

function requestConversion() {
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;
	
	let url = o.UIBasePath + 'Conversion/Conversion?outputType=' + $('#saveAs').val() ;
	
	request(url, data);
}
function requestMetadata(data) {
	

	var response = data.split('|');
	if (response.length > 0) {
		let url = o.UIBasePath + 'api/AsposeWordsMetadata/properties?folderName=' + response[2] + "&fileName=" + response[1];
		$.ajax({
			type: 'POST',
			url: url,			
			contentType: 'application/json',
			cache: false,
			timeout: 600000,
			success: (d) => {
				$.metadata(d, response[2], response[1]);
			},
			error: (err) => {
				if (err.data !== undefined && err.data.Status !== undefined)
					showAlert(err.data.Status);
				else
					showAlert("Error " + err.status + ": " + err.statusText);
			}
		});
	}
}
function requestRedaction() {
	if (!validateSearch())
		return;
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;
	let url = o.UIBasePath + 'Redaction/Redaction?outputType=' + $('#saveAs').val() +
		'&searchQuery=' + encodeURI($('#searchQuery').val()) +
		'&replaceText=' + encodeURI($('#replaceText').val()) +
		'&caseSensitive=' + $('#caseSensitive').prop('checked') +
		'&text=' + $('#text').prop('checked') +
		'&comments=' + $('#comments').prop('checked') +
		'&metadata=' + $('#metadata').prop('checked');
	request(url, data);
}
function validateSearch() {
	if ($("#searchQuery").val().length)
		return true;
	showAlert(o.validationSearchMessage);
	return false;
}
function requestSearch() {
	if (!validateSearch())
		return;
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;
	let url = o.UIBasePath + 'Search/Search?query=' + encodeURI($('#searchQuery').val());
	request(url, data);
}
function validateUnlock() {
	if ($("#passw").val().length)
		return true;
	showAlert(o.validationMessage);
	return false;
}
function requestUnlock() {
	if (!validateUnlock())
		return;
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;
	let url = o.UIBasePath + 'Unlock/Unlock?passw=' + encodeURI($('#passw').val())
		+ '&outputType=' + $('#saveAs').val();
	request(url, data);
}
function requestProtect() {
	if (!validateUnlock())
		return;
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;
	let url = o.UIBasePath + 'Protect/Protect?passw=' + encodeURI($('#passw').val());
		
	request(url, data);
}
function validateComparison() {
	if (fileDrop.droppedFiles.length === 1 && fileDrop.droppedFiles.length === 1)
		return true;
	showAlert(o.FileSelectMessage);
	return false;
}
function requestComparison() {
	if (!validateComparison())
		return;
	let data = fileDrop.prepareFormData();
	if (data === null)
		return;	
	let data2 = fileDrop2.prepareFormData();
	if (data2 === null)
		return;
	for (var entry of data2.entries())
		data.append(entry[0], entry[1]);
	let url = o.UIBasePath + 'Comparison/Comparison';
	request(url, data);
}

function requestViewer(data) {
	
	var url = generateViewerLink(data);
	openIframe(url, '/words/viewer', '/words/view');
}
function requestEditor(data) {
	var url = generateEditorLink(data);
	openIframe(url, '/words/editor', '/words/edit');
}
function prepareDownloadUrl() {
	o.AppDownloadURL = o.AppURL;
	var pos = o.AppDownloadURL.indexOf(':');
	if (pos > 0) 
		o.AppDownloadURL = (pos > 0 ? o.AppDownloadURL.substring(pos + 3) : o.AppURL) + '/download';
	pos = o.AppDownloadURL.indexOf('/');
	o.AppDownloadURL = o.AppDownloadURL.substring(pos);
}

function checkReturnFromViewer() {
	var query = window.location.search;
	if (query.length > 0) {
		o.ReturnFromViewer = true;
		var data = {
			StatusCode: 200,
			FolderName: $.QueryString['id'],
			FileName: $.QueryString['FileName'],
			FileProcessingErrorCode: 0
		};
		var beforeQueryString = window.location.href.split("?")[0];
		window.history.pushState({}, document.title, beforeQueryString);
		if (!o.UploadAndRedirect)
			workSuccess(data);
	}
}

function getInputType() {
    var defaultType = 'html';
    var pathUrl = window.location.pathname.toLowerCase();
    var conversionPos = pathUrl.indexOf('conversion');
    if (conversionPos < 0) {
        return defaultType;
    }
    var conv = pathUrl.substring(conversionPos + 11);
    if (conv.length === 0) {
        return defaultType;
    }
    var arr = conv.split('-');
    console.log(arr[0]);
    return arr[0];
}

$(document).ready(function () {
	prepareDownloadUrl();
	checkReturnFromViewer();
	fileDrop = $('form#UploadFile').filedrop(Object.assign({
		showAlert: showAlert,
		hideAlert: hideAlert,
		showLoader: showLoader,
		progress: progress
	}, o));
	if (o.AppName === "Comparison") {
		fileDrop2 = $('form#UploadFile').filedrop(Object.assign({
			showAlert: showAlert,
			hideAlert: hideAlert,
			showLoader: showLoader,
			progress: progress
		}, o));
	}
	// close iframe if it was opened	
	window.onpopstate = function (event) {
		if ($('div#iframe-wrap').length > 0) {
			closeIframe();
		}
	};

	if (!o.UploadAndRedirect) {
		$('#uploadButton').on('click', o.Method);
	}

	
});
