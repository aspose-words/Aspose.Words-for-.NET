/**
 * GroupDocs.Viewer.JS
 * Copyright (c) 2001-2018 Aspose Pty Ltd
 * Licensed under MIT
 * @author Aspose Pty Ltd
 * @version 1.8.0
 30/07/2019 by Dmitry Zotov
 */

/*
******************************************************************
******************************************************************
GLOBAL VARIABLES
******************************************************************
******************************************************************
*/
var applicationPath;
var preloadPageCount;
var currentDirectory;
var documentGuid;
var folderName;
var documentData = {};
var password = '';
var rewrite;
var loadedPagesCount = 0;
var map = {};
var htmlMode = false;
var htmlModeThumbnails = false;
var thumbnails = false;
var saveRotateState = true;
var enableRightClick = true;
var upload = true;
var printAllowed = true;

var totalPageNumber = 0;
var downloadingPages = [];
var scrollOff = false;

var navigationPane = null;

// add supported formats
map['folder'] = { 'format': '', 'icon': 'fa-folder' };
map['pdf'] = { 'format': 'Portable Document Format', 'icon': 'fa-file-pdf-o' };
map['doc'] = { 'format': 'Microsoft Word', 'icon': 'fa-file-word-o' };
map['docx'] = { 'format': 'Microsoft Word', 'icon': 'fa-file-word-o' };
map['docm'] = { 'format': 'Microsoft Word', 'icon': 'fa-file-word-o' };
map['dot'] = { 'format': 'Microsoft Word', 'icon': 'fa-file-word-o' };
map['dotx'] = { 'format': 'Microsoft Word', 'icon': 'fa-file-word-o' };
map['dotm'] = { 'format': 'Microsoft Word', 'icon': 'fa-file-word-o' };
map['xls'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['xltx'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['xltm'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['tsv'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['xls2003'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['xlsx'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['xlsm'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['xlsb'] = { 'format': 'Microsoft Excel', 'icon': 'fa-file-excel-o' };
map['ppt'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['pptx'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['potx'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['potm'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['pptm'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['ppsm'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['pps'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['ppsx'] = { 'format': 'Microsoft PowerPoint', 'icon': 'fa-file-powerpoint-o' };
map['vsd'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vdx'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vss'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vsx'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vst'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vtx'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vsdx'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vdw'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vstx'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['vssx'] = { 'format': 'Microsoft Visio', 'icon': 'fa-file-code-o' };
map['mpp'] = { 'format': 'Microsoft Project', 'icon': 'fa-file-text' };
map['mpt'] = { 'format': 'Microsoft Project', 'icon': 'fa-file-text' };
map['msg'] = { 'format': 'Microsoft Outlook', 'icon': 'fa-file-text-o' };
map['eml'] = { 'format': 'Microsoft Outlook', 'icon': 'fa-file-text-o' };
map['emlx'] = { 'format': 'Microsoft Outlook', 'icon': 'fa-file-text-o' };
map['one'] = { 'format': 'Microsoft OneNote', 'icon': 'fa-file-word-o' };
map['odt'] = { 'format': 'Open Document Text', 'icon': 'fa-file-word-o' };
map['ott'] = { 'format': 'Open Document Text Template', 'icon': 'fa-file-word-o' };
map['ods'] = { 'format': 'Open Document Spreadsheet', 'icon': 'fa-file-excel-o' };
map['odp'] = { 'format': 'Open Document Presentation', 'icon': 'fa-file-powerpoint-o' };
map['otp'] = { 'format': 'Open Document Presentation', 'icon': 'fa-file-powerpoint-o' };
map['ots'] = { 'format': 'Open Document Presentation', 'icon': 'fa-file-powerpoint-o' };
map['rtf'] = { 'format': 'Rich Text Format', 'icon': 'fa-file-text-o' };
map['txt'] = { 'format': 'Plain Text File', 'icon': 'fa-file-text-o' };
map['csv'] = { 'format': 'Comma-Separated Values', 'icon': 'fa-file-excel-o' };
map['html'] = { 'format': 'HyperText Markup Language', 'icon': 'fa-file-word-o' };
map['mht'] = { 'format': 'HyperText Markup Language', 'icon': 'fa-file-word-o' };
map['mhtml'] = { 'format': 'HyperText Markup Language', 'icon': 'fa-file-word-o' };
map['xml'] = { 'format': 'Extensible Markup Language', 'icon': 'fa-file-word-o' };
map['xps'] = { 'format': 'XML Paper Specification', 'icon': 'fa-file-word-o' };
map['dxf'] = { 'format': 'AutoCAD Drawing File Format', 'icon': 'fa-file-image-o' };
map['dwg'] = { 'format': 'AutoCAD Drawing File Format', 'icon': 'fa-file-image-o' };
map['bmp'] = { 'format': 'Bitmap Picture', 'icon': 'fa-file-image-o' };
map['gif'] = { 'format': 'Graphics Interchange Format', 'icon': 'fa-file-image-o' };
map['jpg'] = { 'format': 'Joint Photographic Experts Group', 'icon': 'fa-file-image-o' };
map['psd'] = { 'format': 'Layered Image File', 'icon': 'fa-file-image-o' };
map['svg'] = { 'format': 'Scalable Vector Graphics', 'icon': 'fa-file-image-o' };
map['jpe'] = { 'format': 'Joint Photographic Experts Group', 'icon': 'fa-file-image-o' };
map['jp2'] = { 'format': 'Compressed Bitmap Image', 'icon': 'fa-file-image-o' };
map['jpeg'] = { 'format': 'Joint Photographic Experts Group', 'icon': 'fa-file-image-o' };
map['jfif'] = { 'format': 'Joint Photographic Experts Group', 'icon': 'fa-file-image-o' };
map['png'] = { 'format': 'Portable Network Graphics', 'icon': 'fa-file-image-o' };
map['tiff'] = { 'format': 'Tagged Image File Format', 'icon': 'fa-file-photo-o' };
map['tif'] = { 'format': 'Tagged Image File Format', 'icon': 'fa-file-photo-o' };
map['epub'] = { 'format': 'Electronic Publication', 'icon': 'fa-file-pdf-o' };
map['ico'] = { 'format': 'Windows Icon', 'icon': 'fa-file-image-o' };
map['webp'] = { 'format': 'Compressed Image', 'icon': 'fa-file-image-o' };
map['mobi'] = { 'format': 'Mobipocket eBook', 'icon': 'fa-file-pdf-o' };
map['tex'] = { 'format': 'LaTeX Source Document', 'icon': 'fa-file-pdf-o' };
map['djvu'] = { 'format': 'Multi-Layer Raster Image', 'icon': 'fa-file-text' };
map['unknown'] = { 'format': 'This format is not supported', 'icon': 'fa-file-o' };

$(document).ready(function () {

	/*
	******************************************************************
	NAV BAR CONTROLS
	******************************************************************
	*/
	if (!enableRightClick) {
		$(document).bind("contextmenu", function (e) {
			e.preventDefault();
		});
	}
	//////////////////////////////////////////////////
	// Toggle navigation dropdown menus
	//////////////////////////////////////////////////
	$('.gd-nav-toggle').on('click', function (e) {
		if ($(this).hasClass('open')) {
			$(this).removeClass('open');
		} else {
			$(this).addClass('open');
		}
		var nav_drop = getElementByClass($(this), '.gd-nav-dropdown');
		toggleNavDropdown(nav_drop);
		//set focus to search input
		$('#gd-search-input').focus();
	});

	//////////////////////////////////////////////////
	// Prevent toggle events on search container click
	//////////////////////////////////////////////////
	$('#gd-nav-search-container').on('click', function (e) {
		e.stopPropagation();
	});

	//////////////////////////////////////////////////
	// Close modal dialog event
	//////////////////////////////////////////////////
	$('.gd-modal-close-action').on('click', closeModal);

	//////////////////////////////////////////////////
	// File or directory click event from file tree
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', '.gd-filetree-name', function (e) {
		var isDir = $(this).parent().find('.fa-folder').hasClass('fa-folder');
		if (isDir) {
			// if directory -> browse
			if (currentDirectory.length > 0) {
				currentDirectory = currentDirectory + "/" + $(this).text();
			} else {
				currentDirectory = $(this).text();
			}
			toggleModalDialog(false, '');
			loadFileTree(currentDirectory);
		} else {
			// if document -> open
			(!$('#gd-btn-zoom-value').hasClass("disabled")) ? $('#gd-btn-zoom-value').addClass("disabled") : "";
			$('#gd-btn-zoom-value > li').unbind("click");
			(!$('#gd-btn-zoom-in').hasClass("disabled")) ? $('#gd-btn-zoom-in').addClass("disabled") : "";
			$('#gd-btn-zoom-in').unbind("click");
			(!$('#gd-btn-zoom-out').hasClass("disabled")) ? $('#gd-btn-zoom-out').addClass("disabled") : "";
			$('#gd-btn-zoom-out').unbind("click");
			(!$('#gd-btn-print').hasClass("disabled")) ? $('#gd-btn-print').addClass("disabled") : "";
			$('#gd-btn-print').unbind("click");
			(!$('#gd-btn-download').hasClass("disabled")) ? $('#gd-btn-download').addClass("disabled") : "";
			clearPageContents();
			toggleModalDialog(false, '');
			password = "";
			documentGuid = $(this).attr('data-guid');
			loadDocument(function (data) {
				generatePagesTemplate(data);
			});
		}
	});

	//////////////////////////////////////////////////
	// Go to parent directory event from file tree
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', '.gd-go-up', function (e) {
		if (currentDirectory.length > 0 && currentDirectory.indexOf('/') == -1) {
			currentDirectory = '';
		} else {
			currentDirectory = currentDirectory.replace(/\/[^\/]+\/?$/, '');
		}
		loadFileTree(currentDirectory);
	});

	//////////////////////////////////////////////////
	// Page navigation event
	//////////////////////////////////////////////////
	$('.gd-nav-btn-pages').on('click', function (e) {
		var pagesAttr = $('#gd-page-num').text().split('/');
		var currentPageNumber = parseInt(pagesAttr[0]);
		var lastPageNumber = parseInt(pagesAttr[1]);

		// single paged document
		if (lastPageNumber === 1) {
			return;
		}

		var prevPage = false;
		// get clicked id
		switch ($(this).attr('id')) {
			case 'gd-btn-page-first':
				currentPageNumber = 1;
				gaEvent('To First Page');
				break;
			case 'gd-btn-page-prev':
				if (currentPageNumber > 1) {
					currentPageNumber = currentPageNumber - 1;
					prevPage = true;
					gaEvent('To Previous Page');
				}
				break;
			case 'gd-btn-page-next':
				if (currentPageNumber < lastPageNumber) {
					currentPageNumber = currentPageNumber + 1;
					gaEvent('To Next Page');
				}
				break;
			case 'gd-btn-page-last':
				currentPageNumber = lastPageNumber;
				gaEvent('To Last Page');
				break;
		}
		// scroll to page
		if (currentPageNumber > 0 && currentPageNumber <= lastPageNumber) {
			var scrollCallback = function() {
				scrollToPage(currentPageNumber);
			};
			// set navigation to current page
			setNavigationPageValues(currentPageNumber, lastPageNumber);
			// load next page
			if (preloadPageCount > 0) {
				// check if next page number is not bigger than total page number
				if (currentPageNumber + 1 <= lastPageNumber) {
					if (prevPage) {
						// load previous page
						// to set correct page size we use global array documentData which contains all info about current document
						appendHtmlContent(currentPageNumber, documentGuid, scrollCallback);
					} else {
						// load next page
						appendHtmlContent(currentPageNumber + 1, documentGuid, scrollCallback);
					}
				} else {
					// load last page if to jump to it via last page button
					appendHtmlContent(currentPageNumber, documentGuid, scrollCallback);
					appendHtmlContent(currentPageNumber - 1, documentGuid);
				}
			}
			//scrollToPage(currentPageNumber);
		}
	});

	//////////////////////////////////////////////////
	// Page scrolling event
	//////////////////////////////////////////////////
	var previousScroll = 0;
	$('#gd-pages').scroll(function () {
		if (scrollOff || downloadingPages.length > 0)
			return;
		
		var pagesAttr = $('#gd-page-num').text().split('/');
		// get current page number
		var currentPageNumber = parseInt(pagesAttr[0]);
		// get last page number
		var lastPageNumber = parseInt(pagesAttr[1]);
		// get scroll direction
		var scrollDown = true;
		var currentScroll = $(this).scrollTop();
		if (currentScroll < previousScroll) {
			scrollDown = false;
		}
		// set scroll direction
		previousScroll = currentScroll;
		var zoom = parseInt($("#gd-zoom-value").html()) / 100;
		var delta = 0.5;
		if (zoom < 1) {
			delta = 1;
		} else {
			delta = delta / zoom;
		}
		for (i = 1; i <= lastPageNumber; i++) {
			// check if page is visible in the view port more than 50%
			if ($('#gd-page-' + i).isOnScreen(delta, delta)) {
				// change current page value
				if (i != currentPageNumber) {
					// set current page number
					setNavigationPageValues(i, lastPageNumber);
				}
				// load next page
				// to set correct page size we use global array documentData which contains all info about current document
				if (preloadPageCount > 0) {
					// if scroll down load next page
					if (scrollDown) {
						if (i + 1 <= lastPageNumber) {
							appendHtmlContent(i + 1, documentGuid);
						} else if (i == lastPageNumber) {
							appendHtmlContent(i, documentGuid);
						}
					} else {
						// if scroll up load previous page
						if (currentPageNumber - 1 >= 1) {
							appendHtmlContent(currentPageNumber - 1, documentGuid);
						}
					}
				}
			}
		}
		if ($(this).scrollTop() == 0 && !scrollDown) {
			setNavigationPageValues(1, lastPageNumber);
		}
	});

	//////////////////////////////////////////////////
	// Clear search input
	//////////////////////////////////////////////////
	$('#gd-nav-search-cancel').on('click', function (e) {
		clearSearch();
		$(this).parent().parent().click();
		gaEvent('Search Clear');
	});

	//////////////////////////////////////////////////
	// Read search input value on input change event
	//////////////////////////////////////////////////
	var search_position = 0;

	var searchTimeout = null;

	var searchHighlight = function(searchQuery) {
		// fix search position if first match
		search_position = 0;
		highlightSearch(searchQuery);
		gaEvent('Search Input');
	};

	var searchFunction = function() {
		const searchQuery = $(this).val();
		if (searchQuery.toString().length > 0 && !searchQuery.match(/^\s+$/)) {
			const data = {
				fileName: documentGuid,
				folderName,
				searchQuery
			};
			// show loading spinner
			showPreloader();

			$.ajax({
				type: 'POST',
				url: getApplicationPath('search'),
				data: JSON.stringify(data),
				contentType: 'application/json',
	
				success: (returnedData) => {
					const promises = [];
					for (var i = 0; i < returnedData.length; i++) {
						if (downloadingPages.indexOf(returnedData[i]) === -1) { // required page was not loaded
							promises.push(new Promise(resolve => {
								appendHtmlContent(returnedData[i], documentGuid, () => {
									resolve();
								});
							}));
						}
					}
					if (promises.length > 0) {
						Promise.all(promises).then(() => {
							// hide loading spinner
							hidePreloader(1500);
							searchHighlight(searchQuery);
						});
					} else {
						searchHighlight(searchQuery);
					}
				},
				error: (xhr, status, error) => {
					// hide loading spinner
					hidePreloader();

					var err = eval('(' + xhr.responseText + ')');
					console.log(err.Message);
					// open error popup
					printMessage(err.message);
				}
			});
		}
	};

	// user has pressed a key in the search input
	$('#gd-search-input').on('keydown', function(e) {
		if (searchTimeout !== null) {
			window.clearTimeout(searchTimeout);
			searchTimeout = null;
		}
	});

	// user has released a key in the search input
	$('#gd-search-input').on('keyup', function(e) {
		searchTimeout = window.setTimeout(searchFunction.bind(this), 1000); // 1 sec. delay before ajax request
	});

	//////////////////////////////////////////////////
	// Search next event
	//////////////////////////////////////////////////
	$('#gd-nav-search-next').on('click', function (e) {
		var count = 0;
		var zoomValue = getZoomValue();
		// remove/clear previously selected highlights
		$('#gd-pages').find('.gd-highlight-select').removeClass('gd-highlight-select');
		// search for matched elements
		$('.gd-highlight').each(function (e) {
			if (count == search_position) {
				// add selected highlight
				$(this).addClass('gd-highlight-select');
				// scroll to next page
				$('#gd-pages').scrollTo(this, {
					offsetTop: 150,
					zoom: zoomValue
				});
				// check if this is last rearch result instance
				if (search_position >= getTotalSearchMatches()) {
					// deactivate next button
				} else {
					// increment search position
					search_position++;
				}
				// end each loop
				return false;
			}
			count++;
		});
		setSearchMatchCount(search_position, getTotalSearchMatches());
		gaEvent('Search Next');
	});

	//////////////////////////////////////////////////
	// Search prev event
	//////////////////////////////////////////////////
	$('#gd-nav-search-prev').on('click', function (e) {
		var count = 1;
		var prev;
		var zoomValue = getZoomValue();
		// remove/clear previously selected highlights
		$('#gd-pages').find('.gd-highlight-select').removeClass('gd-highlight-select');
		// search for matched elements
		$('.gd-highlight').each(function (e) {
			if ((count == (search_position)) && (prev != undefined)) {
				// add selected highlight
				$(prev).addClass('gd-highlight-select');
				// scroll to previous page
				$('#gd-pages').scrollTo(prev, {
					offsetTop: 150,
					zoom: zoomValue
				});
				// check if this is first search result instance
				if (search_position <= 1) {
					// deactivate prev button
				} else {
					// decrement search position
					search_position--;
				}
				// end each loop
				return false;
			}
			count++;
			// store last instance
			prev = $(this);
		});
		setSearchMatchCount(search_position, getTotalSearchMatches());
		gaEvent('Search Prev');
	});

	//////////////////////////////////////////////////
	// Open/Close thumbnails event
	//////////////////////////////////////////////////
	$('#gd-nav-right').on('click', function () {
		// open/close sidebar
		$('#gd-thumbnails').toggleClass('active');
		gaEvent('Thumbnails Open');
	});

	//////////////////////////////////////////////////
	// Thumbnail click event
	//////////////////////////////////////////////////
	$('#gd-thumbnails-panzoom').on('click', '.gd-page', function () {
		// get clicked thumbnail page number
		var page = parseInt($(this).attr('id').split('-')[3]);
		var pagesAttr = $('#gd-page-num').text().split('/');
		// get last page number
		var lastPageNumber = parseInt(pagesAttr[1]);

		if (preloadPageCount > 0) {
			if (page == lastPageNumber) {
				appendHtmlContent(page, documentGuid);
				appendHtmlContent(page, documentGuid);
			} else {
				appendHtmlContent(page, documentGuid);
				appendHtmlContent(page + 1, documentGuid);
			}
		}
		// set navigation to current page
		setNavigationPageValues(page, lastPageNumber);
		scrollToPage(page);
		gaEvent('Thumbnails Click');
	});

	//////////////////////////////////////////////////
	// Rotate counterclockwise event
	//////////////////////////////////////////////////
	$('#gd-btn-counterclockwise').on('click', function (e) {
		rotatePages("-90");
		gaEvent('Rotate CCW');
	});

	//////////////////////////////////////////////////
	// Rotate clockwise event
	//////////////////////////////////////////////////
	$('#gd-btn-clockwise').on('click', function (e) {
		rotatePages("90");
		gaEvent('Rotate CW');
	});

	//////////////////////////////////////////////////
	// Download event
	//////////////////////////////////////////////////
	$('#gd-btn-download').on('click', function (e) {
		downloadDocument();
		gaEvent('Download');
	});

	//////////////////////////////////////////////////
	// Select files for upload event
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('change', '#gd-upload-input', function (e) {
		// get selected files
		$("#gd-upload-input-checkbox").prop("checked", false);
		var input = $(this);
		$.each(input.get(0).files, function (index, file) {
			uploadDocument(file);
		});
	});

	//////////////////////////////////////////////////
	// Select file from HDD icon click
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', '.gd-hdd', function (e) {
		$("#gd-upload-input").click();
	});

	//////////////////////////////////////////////////
	// Open URL input event
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', '#gd-upload-url', function () {
		$('#gd-url-wrap').show();
	});

	//////////////////////////////////////////////////
	// Add file via URL event
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', '#gd-add-url', function () {
		$('#gd-url-wrap').hide();
		uploadDocument(null, $("#gd-url").val());
		$('#gd-url').val('');
	});

	//////////////////////////////////////////////////
	// Open document button (upload dialog) click
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', '#gd-open-document', function (e) {
		toggleModalDialog(false, '');
		loadFileTree('');
	});

	//////////////////////////////////////////////////
	// Open search event
	//////////////////////////////////////////////////
	$('#gd-btn-search').on('click', function (e) {
		if ($('#gd-nav-search-container').parent().find('span').is(':visible')) {
			$('#gd-nav-search-container').parent().find('span').css('display', 'none');
		} else {
			$('#gd-nav-search-container').parent().find('span').css('display', 'initial');
		}
		gaEvent('Search Open');
	});

	//////////////////////////////////////////////////
	// Submit password button click (password required modal)
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', "#gd-password-submit", function (e) {
		password = $('#gd-password-input').val();
		$('#gd-password-input').val('');
		toggleModalDialog(false, '');
		loadDocument(function (data) {
			generatePagesTemplate(data);
		});
	});

	//////////////////////////////////////////////////
	// Click on modal body event (used to change slide in swiper)
	//////////////////////////////////////////////////
	$('.gd-modal-body').on('click', '#gd-modal-content', function (e) {
		if (isMobile()) {
			if ($('#gd-upload-files-table > div').length > 0) {
				var swiper = new Swiper('.swiper-container');
				if (typeof swiper.length == 'undefined') {
					swiper.slideNext();
					swiper.slidePrev();
				}
				for (var i = 0; i < swiper.length; i++) {
					swiper[i].slideNext();
					swiper[i].slidePrev();
				}
			}
		}
	});

	$('#modalDialog').on('click', function (event) {
		if ($(event.target).hasClass("fas fa-chevron-down")) {
			event.preventDefault();
			$("#gd-upload-input-checkbox").prop("checked", true);
			return;
		}
		if (event.target.tagName != "LABEL" && event.target.tagName != "LI" && event.target.tagName != "INPUT") {
			$("#gd-upload-input-checkbox").prop("checked", false);
		}
	});

	//
	// END of document ready function
});

/*
******************************************************************
FUNCTIONS
******************************************************************
*/

/**
* Load file tree
* @param {string} dir - files location directory
*/
function loadFileTree(dir, multiple) {
	if (!multiple) {
		multiple = false;
	}
	var data = { path: dir };
	currentDirectory = dir;
	// clear previously entered password
	clearPassword();
	// show loading spinner
	$('#gd-modal-spinner').show();
	// get data
	$.ajax({
		type: 'POST',
		url: getApplicationPath('loadFileTree'),
		data: JSON.stringify(data),
		contentType: 'application/json',
		success: function (returnedData) {
			if (returnedData.message != undefined) {
				// open error popup
				printMessage(returnedData.message);
				return;
			}
			// assembly modal html
			$('.gd-modal-body').html(''); // clear previous data
			toggleModalDialog(true, "Open Document", getHtmlFileBrowser(multiple));
			initDragNDrop(uploadDocument);
			// hide loading spinner
			$('#gd-modal-spinner').hide();
			// append files to tree list
			$.each(returnedData, function (index, elem) {
				// document name
				var name = elem.name;
				// document guid
				var guid = elem.guid;
				// document size
				var size = elem.size;
				// convert to proper size
				var new_size = size + ' Bytes';
				if ((size / 1024 / 1024) > 1) {
					new_size = (Math.round((size / 1024 / 1024) * 100) / 100) + ' MB';
				} else if ((size / 1024) > 1) {
					new_size = (Math.round((size / 1024) * 100) / 100) + ' KB';
				}
				// document format
				var docFormat = (getDocumentFormat(name, elem.isDirectory) == undefined) ? 'fa-folder' : getDocumentFormat(name, elem.isDirectory);
				var folderClass = (docFormat.format == "") ? "gd-folder-name" : "";
				var checkBoxes = "";
				var places = "";
				var single = "single";
				if (multiple) {
					if (elem.isDirectory) {
						checkBoxes = '<div class="gd-file-checkbox empty"></div>';
					} else {
						checkBoxes = '<div class="gd-file-checkbox"><input type="checkbox" id="' + name + '" name="' + name + '" class="gd-checkbox"></div>';
					}
					places = '<div><i class="fas fa-plus"></i></div>';
					single = "";
				}
				// append document
				$('.gd-modal-table-body').append(
					'<div class="gd-file-table-item">' +
					checkBoxes +
					'<div class="gd-filetree-name ' + single + '" data-guid="' + guid + '">' +
					'<i class="fa ' + docFormat.icon + '"></i>' +
					'<div class="gd-file-name ' + folderClass + '">' + name +
					'<div class="gd-file-format">' + docFormat.format + '</div>' +
					'</div>' +
					'</div >' +
					'<div class="gd-file-size">' + new_size + '</div>' +
					places +
					'</div>');
			});
		},
		error: function (xhr, status, error) {
			var err = eval("(" + xhr.responseText + ")");
			console.log(err.Message);
			// hide loading spinner
			$('#gd-modal-spinner').hide();
			// open error popup
			printMessage(err.message);
		}
	});
}

function fadeAll(on) {
	if (on) {
		$('#gd-container-fade').show();
		//$('#gd-container-fade-text').show();
		$('#page-loading').show();
		$('#htmlloader').show();
	} else {
		$('#gd-container-fade').hide();
		//$('#gd-container-fade-text').hide();
		$('#page-loading').fadeOut(600);
		$('#htmlloader').hide();
	}
}

function showPreloader(timeout) {
	window.setTimeout(() => {
		$('#preloader').removeClass('hidden');
	}, timeout ? timeout : 0);
}

function hidePreloader(timeout) {
	window.setTimeout(() => {
		$('#preloader').addClass('hidden');
	}, timeout ? timeout : 0);
}

function toggleThumbnails() {
	$('#gd-thumbnails').toggleClass('active');
}

function clickHeader(page, name) {
	const p1 = new Promise(resolve => {
		if (page > 1) {
			appendHtmlContent(page - 1, documentGuid, () => {
				resolve();
			});
		} else {
			resolve();
		}
	});
	const p2 = new Promise(resolve => {
		showPreloader();
		appendHtmlContent(page, documentGuid, () => {
			scrollToPage(page, false, scrollParams => {
				if (name) {
					var ank = '#gd-page-' + page + ' span:contains(\'' + name + '\')';
					var hdr = $(ank);
					if (hdr.length > 0) {
						$('#gd-pages').scrollTo(ank, scrollParams);
					}
				}
			});
			hidePreloader();
			resolve();
		})
	});
	const p3 = new Promise(resolve => {
		if (page < totalPageNumber) {
			appendHtmlContent(page + 1, documentGuid, () => {
				resolve();
			});
		} else {
			resolve();
		}
	});

	Promise.all([p1, p2, p3]);
}

/**
* Open/Load document
* @param {object} callback - document pages array
*/
function loadDocument(callback) {
	// clear global documentData array from previous document info
	documentData = {};
	loadedPagesCount = 0;
	// fade all controls
	fadeAll(true);
	// get document description
	var data = {
		fileName: documentGuid,
		folderName: folderName,
		password: password
	};
	$.ajax({
		type: 'POST',
		url: getApplicationPath('documentInfo'),
		data: JSON.stringify(data),
		global: false,
		contentType: "application/json",
		success: function (returnedData) {
			if (returnedData.message != undefined) {
				fadeAll(false);
				if (returnedData.message == "Password Required") {
					openPasswordModal();
				} else if (returnedData.message == "Incorrect password") {
					openPasswordModal(returnedData.message);
					return;
				} else {
					// open error popup
					printMessage(returnedData.message);
				}
				return;
			}
			if (returnedData.navigationPane && returnedData.navigationPane.length > 0) {
				const STYLES = {
					1  : 'header-1',
					2  : 'header-2',
					3  : 'header-3',
					62 : 'title',
					74 : 'subtitle'
				};
				navigationPane = returnedData.navigationPane;
				var ul = $('<ul class="navigation" />');
				for (var i = 0; i < navigationPane.length; i++) {
					if (navigationPane[i].name) {
						var li = $('<li />');
						li.addClass(STYLES[navigationPane[i].style]);
						li.append($('<a href="javascript:void(0)" onclick="clickHeader(' + navigationPane[i].pageNumber + ', \'' + navigationPane[i].name.trim() + '\')">' + navigationPane[i].name + '</a>'));
						ul.append(li);
					}
				}
				$('#gd-thumbnails-headers').append(ul);

				// open the navigation pane instantly
				$('#gd-thumbnails').toggleClass('active');
			} else {
				$('#gd-thumbnails ul.tabs').addClass('hidden');
			}
			// set current document data globally
			documentData = returnedData;
			// get total page number
			totalPageNumber = documentData.pages.length;
			// single paged document
			if (totalPageNumber === 1) {
				$('.gd-nav-btn-pages').addClass('disabled');
				$('#gd-page-num').addClass('disabled');
			}
			// set total page number on navigation panel
			setNavigationPageValues('1', totalPageNumber);
		},
		error: function (xhr, status, error) {
			fadeAll(false);
			var err = new Error(xhr.responseText ? xhr.responseText : (error ? error : 'Unknown error'));
			console.log(err.message);
			// open error popup
			printMessage(err.message);
			documentGuid = '';
		}
	}).done(function (data) {
		// return POST data
		if (data.message == undefined && callback) {
			callback(data);
		}
	});
}

function loadThumbnails() {
	var data = {
		fileName: documentGuid,
		folderName: folderName,
		password: password
	};
	$.ajax({
		type: 'POST',
		url: getApplicationPath('thumbnails'),
		data: JSON.stringify(data),
		global: false,
		contentType: "application/json",
		success: function (returnedData) {
			if (returnedData.message != undefined) {
				console.log(returnedData.message);
				return;
			}
			$.each(returnedData, function (index, elem) {
				var pageNumber = elem.number;
				renderThumbnails(pageNumber, elem);
			});
			//$('#gd-thumbnails').addClass('active');
		},
		error: function (xhr, status, error) {
			var err = eval("(" + xhr.responseText + ")");
			console.log(err ? err.Message : error);
		}
	});
}

function loadPrint() {
	var data = {
		fileName: documentGuid,
		folderName: folderName,
		password: password
	};
	if (preloadPageCount != 0 && document.location.pathname.indexOf("view") >= 0) {
		if (getDocumentFormat(documentGuid).format == "Portable Document Format") {
			printPdf();
		} else {
			$.ajax({
				type: 'POST',
				url: getApplicationPath('print'),
				data: JSON.stringify(data),
				global: false,
				contentType: "application/json",
				success: function (returnedData) {
					if (returnedData.message != undefined) {
						console.log(returnedData.message);
						return;
					}
					var pagesHtml = "";
					$.each(returnedData, function (index, elem) {
						var data = '';
						if (htmlMode) {
							data = elem.data;
						} else {
							data = '<image style="width: inherit !important" class="gd-page-image" src="data:image/png;base64,' + elem.data + '" alt></image>';
						}
						pagesHtml = pagesHtml + '<div id="gd-page-' + elem.number + '" class="gd-page" style="min-width: ' +
							elem.width + 'px; min-height: ' + elem.height + 'px;">' +
							'<div class="gd-wrapper">' + data + '</div>' +
							'</div>';
					});
					renderPrint(pagesHtml);
				},
				error: function (xhr, status, error) {
					var err = eval("(" + xhr.responseText + ")");
					console.log(err ? err.Message : error);
				}
			});
		}
	} else {
		renderPrint();
	}
}

function renderPrint(pages) {
	// force each document page to be printed as a new page
	var cssPrint = '<style>' +
		'.gd-page {height:100% !important; page-break-after:always; page-break-inside:avoid;} .gd-page:last-child {page-break-after:avoid; overflow:hidden;}';
	cssPrint = cssPrint + '</style>';
	var windowObject = window.open('', documentGuid, "width=750,height=650,top=50,left=50,toolbars=yes,scrollbars=yes,status=yes,resizable=yes");

	// windowObject.focus();
	windowObject.document.head.innerHTML = cssPrint;

	// add current document into the print window
	if (!pages && pages != "") {
	// get current document content
		var documentContainer = $("#gd-panzoom");
		// set correct page orientation if page were rotated
		documentContainer.find(".gd-page").each(function (index, page) {
			if ($(page).css("transform") != "none") {
				cssPrint = cssPrint + "#" + $(page).attr("id") + "{transform: rotate(0deg) !important;}";
			}
		});
		pages = documentContainer[0].innerHTML;
	}

	// add current document into the print window
	windowObject.document.body.innerHTML = pages;

	windowObject.document.title = documentGuid;

	window.setTimeout(function() {
		windowObject.focus();
		windowObject.print();
		window.setTimeout(function() {
			windowObject.close();
		}, 0);
	}, 3000);
}

function printPdf() {
	if (documentGuid != "" && typeof documentGuid != "undefined") {
		var data = { guid: documentGuid, password: password };

		var request = new XMLHttpRequest();
		request.open('POST', getApplicationPath('printPdf'), true);
		request.setRequestHeader('Content-Type', 'application/json');
		request.responseType = 'blob';
		request.onload = function () {
			// Only handle status code 200
			if (request.status === 200) {

				var file = new Blob([request.response], { type: 'application/pdf' });
				var fileURL = URL.createObjectURL(file);
				var windowObject = window.open(fileURL, "PrintWindow", "width=750,height=650,top=50,left=50,toolbars=yes,scrollbars=yes,status=yes,resizable=yes");
				windowObject.focus();

				$(windowObject.document).ready(function () {
					windowObject.document.close();
					windowObject.focus();
					windowObject.onafterprint = function (e) {
						$(windowObject).off('mousemove', windowObject.onafterprint);
						windowObject.close();
					};
					windowObject.print();
					setTimeout(function () {
						$(windowObject).on('mousemove', windowObject.onafterprint);
					}, 3000);
				});
			}
		};

		request.send(JSON.stringify(data));
	}
}

function urlExists(url, callback) {
	$.ajax({
		type: 'HEAD',
		url: url,
		success: function () {
			callback(true);
		},
		error: function () {
			callback(false);
		}
	});
}

/**
* Generate empty pages temples before the actual get pages request
* @param {object} data - document pages array
*/
function generatePagesTemplate(data) {
	// hide loading text only
	$('#gd-container-fade-text').hide();
	// loop though pages
	$.each(data.pages, function (index, elem) {
		var pageNumber = elem.number;
		var pageWidth = elem.width;
		var pageHeight = elem.height;
		var pageData = elem.data;
		var overflow = "";
		var overflowThumbnails = "";
		if (getDocumentFormat(documentGuid).format == "Microsoft Excel") {
			pageWidth = 800;
			pageHeight = 450;
			overflow = "overflow: auto";
			overflowThumbnails = "overflow: hidden";
		}
		// append empty page
		$('#gd-panzoom').append(
			'<div id="gd-page-' + pageNumber + '" class="gd-page" style="min-width: ' + pageWidth + 'px; min-height: ' + pageHeight + 'px;' + overflow + '">' +
			'</div>'
		);
		var gd_page = $('#gd-page-' + pageNumber);
		if (pageData == null || (pageData != null && preloadPageCount > 0)) {
			//gd_page.append('<div class="gd-page-spinner"><i class="fa fa-circle-o-notch fa-spin"></i> &nbsp;Loading... Please wait.</div>');
			if (pageNumber <= preloadPageCount) {
				appendHtmlContent(pageNumber, documentGuid);
			}
		} else {
			renderPage(gd_page, elem, documentGuid);
		}
		if (thumbnails) {
			$('#gd-thumbnails-panzoom').append(
				//'<div id="gd-thumbnails-page-' + pageNumber + '" class="gd-page" style="min-width: ' + pageWidth + 'px; min-height: ' + pageHeight + 'px;' + overflowThumbnails + '">' +
				'<div id="gd-thumbnails-page-' + pageNumber + '" class="gd-page" style="' + overflowThumbnails + '">' +
				'</div>'
			);
			if (preloadPageCount == 0) {
				renderThumbnails(pageNumber, elem);
			}
		}
	});
	if (thumbnails && preloadPageCount != 0) {
		loadThumbnails();
	}
}

/**
* Append html content to an empty page
* @param {int} pageNumber - page number
* @param {string} documentName - document name/id
*/
function appendHtmlContent(pageNumber, documentName, callback) {
	// initialize data
	var gd_page = $('#gd-page-' + pageNumber);

	if (!gd_page.hasClass('loaded')) {
		downloadingPages.push(pageNumber);
		gd_page.addClass('loaded');
		// get document description
		var data = {
			fileName: documentGuid,
			folderName: folderName,
			pageNumber: pageNumber,
			password: password
		};
		$.ajax({
			type: 'POST',
			url: getApplicationPath('page'),
			data: JSON.stringify(data),
			global: false,
			contentType: "application/json",
			success: function (htmlData) {
				if (htmlData.error != undefined) {
					// open error popup
					printMessage(htmlData.error);
					return;
				}
				// remove spinner
				//gd_page.find('.gd-page-spinner').hide();
				renderPage(gd_page, htmlData, documentName);
			},
			error: function (xhr, status, error) {
				fadeAll(false);
				var err = eval("(" + xhr.responseText + ")");
				// open error popup
				printMessage(err ? err.message : 'Error occurred while loading');
			}
		}).done(function() {
			if (callback !== undefined)
				callback();
			downloadingPages.pop();
		});
	}
	else if (callback !== undefined)
		callback();
}

function renderThumbnails(pageNumber, pageData) {
	var gd_page = $('#gd-page-' + pageNumber);
	var width = pageData.width;
	var height = pageData.height;
	var zoomValue = gd_page[0].style.zoom;
	// fix thumbnails only when any of document pages is loaded.
	// this is required to fix issue with thumbnails resolution
	isPageLoaded($('#gd-page-1')).then(function (element) {
		var gd_thumbnails_page = $('#gd-thumbnails-page-' + pageNumber);
		if (htmlModeThumbnails) {
			gd_thumbnails_page.append('<div class="gd-wrapper">' + pageData.data + '</div>');
			if (getDocumentFormat(documentGuid).format == "Microsoft PowerPoint") {
				zoomValue = 0.7;
			}
			width = $("#gd-page-1").innerWidth();
			height = $("#gd-page-1").innerHeight();
			if (getDocumentFormat(documentGuid).format == "Microsoft Excel") {
				width = 800;
				height = 450;
				zoomValue = 1;
			}
			gd_thumbnails_page.css('width', width);
			gd_thumbnails_page.css('height', height);
			gd_thumbnails_page.css('zoom', zoomValue);
		} else {
			// if current document if image file fix its zoom
			if (getDocumentFormat(documentGuid).icon.search("image") > 0 || getDocumentFormat(documentGuid).icon.search("photo") > 0) {
				if (width > ($("#gd-thumbnails").width() * 2)) {
					zoomValue = 0.5;
				} else {
					zoomValue = 1.2;
				}
			}
			if (getDocumentFormat(documentGuid).format == "Microsoft Excel") {
				width = 800;
				height = 450;
				zoomValue = 1;
			}
			if (getDocumentFormat(documentGuid).format == "Microsoft PowerPoint") {
				if (width > ($("#gd-thumbnails").width() * 2)) {
					zoomValue = 0.7;
				}
			}
			// set correct size
			gd_thumbnails_page.css('width', width);
			gd_thumbnails_page.css('height', height);
			gd_thumbnails_page.css('zoom', zoomValue);
			// append page image, in image mode append occurred after setting the size to avoid zero size usage
			gd_thumbnails_page.append('<div class="gd-wrapper">' +
				'<image style="width: inherit !important" class="gd-page-image" src="data:image/png;base64,' + pageData.data + '" alt></image>' +
				'</div>');
		}
		// rotate page if it were rotated earlier
		if (pageData.angle != 0) {
			gd_thumbnails_page.css('animation', 'none');
			gd_thumbnails_page.css('transition-property', 'none');
			gd_thumbnails_page.css('transform', 'rotate(' + pageData.angle + 'deg)');
			if (pageData.angle == 90 || pageData.angle == 270) {
				// set styles for HTML mode
				if (htmlModeThumbnails) {
					if (gd_thumbnails_page.width() > gd_thumbnails_page.height()) {
						gd_thumbnails_page.addClass("gd-thumbnails-landscape-rotated");
					} else {
						gd_thumbnails_page.addClass("gd-thumbnails-landscape");
					}
				} else {
					if (gd_thumbnails_page.width() > gd_thumbnails_page.height()) {
						gd_thumbnails_page.addClass("gd-thumbnails-landscape-image-rotated");
					} else {
						gd_thumbnails_page.addClass("gd-thumbnails-landscape-image");
					}
					gd_thumbnails_page.find("img").removeClass("gd-page-image");
				}
			} else {
				gd_thumbnails_page.removeClass("gd-thumbnails-landscape");
				gd_thumbnails_page.removeClass("gd-thumbnails-landscape-image");
			}
		}
	});
}

/**
* Render current document page
* @param {Object} gd_page - current page jQuery object
* @param {Object} pageData - document page
* @param {string} documentName - current document name
* @param {int} pageNumber - current page number
*/
function renderPage(gd_page, pageData, documentName) {
	//console.log('render: ' + pageData.number);
	var width = pageData.width;
	var height = pageData.height;
	// fix zoom in/out scaling
	var zoomValue = 1;
	if (getDocumentFormat(documentGuid).format == "Microsoft Excel") {
		width = 800;
		height = 450;
	}
	// append page content in HTML mode
	if (htmlMode) {
		// append page
		gd_page.append('<div class="gd-wrapper">' + pageData.data + '</div>');
		// fix to avoid using the spinner DIV size
		if (preloadPageCount == 0) {
			var pageSize = setDocumentPageSize(gd_page, width, height);
			width = pageSize.width;
			height = pageSize.height;
		} else {
			// set correct width and height for document pages
			var pageSize = setDocumentPageSize(gd_page, width, height);
			width = pageSize.width;
			height = pageSize.height;
		}
		// check if page is horizontally displayed
		if (width > height || width >= $(window).width()) {
			zoomValue = 0.79;
		}
		// set correct size
		gd_page.css('width', width);
		gd_page.css('height', height);
		if (getDocumentFormat(documentGuid).format == "Microsoft Excel") {
			zoomValue = 1;
		}
		gd_page.css('zoom', zoomValue);
	} else { // render document in image mode
		// check if page is horizontally displayed
		if (width > height) {
			zoomValue = 0.79;
		}
		// if current document if image file fix its zoom
		if (getDocumentFormat(documentGuid).icon.search("image") > 0 || getDocumentFormat(documentGuid).icon.search("photo") > 0) {
			if (width > $(window).width()) {
				zoomValue = 0.79;
			}
		} else {
			zoomValue = 1.2;
		}
		// set correct size
		gd_page.css('width', width);
		gd_page.css('height', height);
		if (getDocumentFormat(documentGuid).format == "Microsoft Excel") {
			zoomValue = 1;
		}
		gd_page.css('zoom', zoomValue);
		// append page image, in image mode append occurred after setting the size to avoid zero size usage
		gd_page.append('<div class="gd-wrapper">' +
			'<image style="width: inherit !important" class="gd-page-image" src="data:image/png;base64,' + pageData.data + '" alt></image>' +
			'</div>');
	}
	// set correct width and high for OneNote format
	if (documentName.substr((documentName.lastIndexOf('.') + 1)) == "one") {
		if (htmlMode) {
			$(".gd-wrapper").css("width", "initial");
		} else {
			$(".gd-wrapper").css("width", "inherit");
		}
	}
	// rotate page if it were rotated earlier
	if (pageData.angle != 0) {
		gd_page.css('animation', 'none');
		gd_page.css('transition-property', 'none');
		gd_page.css('transform', 'rotate(' + pageData.angle + 'deg)');
		if (pageData.angle == 90 || pageData.angle == 270) {
			// set styles for HTML mode
			if (htmlMode) {
				if (gd_page.width() > gd_page.height()) {
					gd_page.addClass("gd-landscape-rotated");
				} else {
					gd_page.addClass("gd-landscape");
				}
			} else {
				// set style for image mode
				if (gd_page.width() > gd_page.height()) {
					gd_page.addClass("gd-landscape-image-rotated");
				} else {
					gd_page.addClass("gd-landscape-image");
				}
			}
		} else {
			gd_page.removeClass("gd-landscape");
			gd_page.removeClass("gd-landscape-image");
		}
	}

	

	
	var pagesAttr = $('#gd-page-num').text().split('/');
	var lastPageNumber = parseInt(pagesAttr[1]);
	if (loadedPagesCount == lastPageNumber || preloadPageCount !== 0 && loadedPagesCount === 0)
	{
		$('#gd-btn-zoom-value > li').bind("click", function () { setZoomLevel.apply(this); });
		$('#gd-btn-zoom-value').removeClass("disabled");
		$('#gd-btn-zoom-in').bind('click', zoomIn);
		$('#gd-btn-zoom-out').bind('click', zoomOut);
		$('#gd-btn-zoom-in').removeClass('disabled');
		$('#gd-btn-zoom-out').removeClass('disabled');
		if (printAllowed || typeof printAllowed == "undefined") {
			$('#gd-btn-print').bind('click', printDocument);
			$('#gd-btn-print').removeClass('disabled');
		}
		$('#gd-btn-download').removeClass('disabled');
	}
	
	if (preloadPageCount === 0 && loadedPagesCount === 0 || 
		loadedPagesCount + 1 === Math.min(preloadPageCount, lastPageNumber)) {
		fadeAll(false);
	}
	loadedPagesCount = loadedPagesCount + 1;
}

/**
* Calculate document page width and height
* @param {Object} gd_page = document page placeholder object;
* @param {int} width - current page width
* @param {int} height - current page height
*/
function setDocumentPageSize(gd_page, width, height) {
	var pageSize = {
		width: 0,
		height: 0
	};
	if (gd_page.innerWidth() >= width - 1 && gd_page.innerHeight() >= height - 1 && !/Edge/.test(navigator.userAgent)) {
		pageSize.width = gd_page.innerWidth();
		pageSize.height = gd_page.innerHeight();
	} else if (gd_page.innerWidth() >= width - 1 && gd_page.innerHeight() >= height - 1 && /Edge/.test(navigator.userAgent)) {
		pageSize.width = gd_page.innerWidth() / 10;
		pageSize.height = gd_page.innerHeight();
		if (getDocumentFormat(documentGuid).format.search("Word") != -1) {
			pageSize.width = gd_page.innerWidth();
		}
	}
	if (getDocumentFormat(documentGuid).format == "Microsoft Excel") {
		pageSize.width = 800;
		pageSize.height = 450;
	}
	return pageSize;
}

/**
* Get document format (type)
* @param {string} filename - document name
* @param {boolean} isDirectory - define if the current element is directory or file
*/
function getDocumentFormat(filename, isDirectory) {
	if (!isDirectory) {
		if (typeof map[filename.split('.').pop().toLowerCase()] == "undefined") {
			if (filename.split('.').length > 1) {
				return map["unknown"];
			} else {
				return map["folder"];
			}
		} else {
			return map[filename.split('.').pop().toLowerCase()];
		}
	} else {
		return map["folder"];
	}
}

/**
* Get application path for GET/POST requests
* @param {string} context - application context
*/
function getApplicationPath(context) {
	if (applicationPath != null) {
		if (applicationPath.slice(-1) == '/') {
			return applicationPath + context;
		} else {
			return applicationPath + "/" + context;
		}
	} else {
		return context;
	}
}

/**
* Search for element by class (recursive)
* @param {object} target - object where to search for an id
* @param {string} class_id - class id
*/
function getElementByClass(target, class_id) {
	var elem = target.find(class_id);
	if (!elem.hasClass(class_id.slice(1))) {
		return getElementByClass(target.parent(), class_id);
	} else {
		return elem;
	}
}

/**
* Toggle modal dialog
* @param {boolean} open - open/close value
* @param {string} title - title to display in modal dialog (popup)
*/
function toggleModalDialog(open, title, content) {
	if (open) {
		$('#modalDialog .gd-modal-title').text(title);
		$('#modalDialog')
			.css('opacity', 0)
			.fadeIn('fast')
			.animate(
				{ opacity: 1 },
				{ queue: false, duration: 'fast' }
			);
		$('#modalDialog').addClass('in');
		$(".gd-modal-body").append(content);
	} else {
		$('#modalDialog').removeClass('in');
		$('#modalDialog')
			.css('opacity', 1)
			.fadeIn('fast')
			.animate(
				{ opacity: 0 },
				{ queue: false, duration: 'fast' }
			)
			.css('display', 'none');
		$(".gd-modal-body").html('');
	}
}

/**
* Toggle top navigation menus (zoom, search)
* @param {object} target - dropdown target to be opened/closed
*/
function toggleNavDropdown(target) {
	var isOpened = target.hasClass('opened');
	if (!isOpened) {
		$(target).addClass('opened');
		$(target)
			.css('opacity', 0)
			.slideDown('fast')
			.animate(
				{ opacity: 1 },
				{ queue: false, duration: 'fast' }
			);
	} else {
		$(target).removeClass('opened');
		$(target)
			.css('opacity', 1)
			.slideUp('fast')
			.animate(
				{ opacity: 0 },
				{ queue: false, duration: 'fast' }
			);
	}
}

/**
* Highlight search results
* @param {string} text - text to search
*/
function highlightSearch(text) {
	clearHighlightSearch();
	if (text.length > 1) {
		var textNodes = $('#gd-pages .gd-wrapper div').find('*').contents().filter(function () {
			return this.nodeType === 3;
		});
		textNodes.each(function () {
			var $this = $(this);
			var content = $this.text();
			// ignor upper/lower cases
			var regEx = new RegExp(text, "ig");
			// find mattching string
			var str_start = content.search(regEx);
			var originalString = 'did not worked!';
			if (str_start >= 0) {
				var originalString = content.substring(str_start, str_start + text.length);
			}
			// add highlight
			content = content.replace(regEx, '<span class="gd-highlight">' + originalString + '</span>');
			// update to new content
			$this.replaceWith(content);
		});
		var totalSearchMatches = getTotalSearchMatches();
		setSearchMatchCount(0, totalSearchMatches);
		if (totalSearchMatches > 0) {
			$('#gd-nav-search-next').click();
		}
	}
}

/**
* Clear search input
*/
function clearSearch() {
	$('#gd-nav-search-container :input').val('');
	setSearchMatchCount(0, 0);
	clearHighlightSearch();
}

/**
* Zoom document
* @param {int} zoom_val - zoom value from 0 to 100
*/
function setZoomValue(zoom_val) {
	// adapt value for css
	var zoom_val_non_webkit = zoom_val / 100;
	var zoom_val_webkit = Math.round(zoom_val) + '%';
	// display zoom value
	setNavigationZoomValues(zoom_val_webkit);
	if (navigator.userAgent.toLowerCase().indexOf('firefox') > -1) {
		if (zoom_val > 100) {
			$(".gd-page").each(function (index, page) {
				if (!page.id.startsWith("gd-thumbnails") && !$(page).hasClass("gd-page-zoomed"))
					$(page).addClass("gd-page-zoomed");
			});
		} else {
			$(".gd-page").each(function (index, page) {
				$(page).removeClass("gd-page-zoomed");
			});
		}
	}
	// set css zoom values
	var style = [
		'zoom: ' + zoom_val_webkit,
		'zoom: ' + zoom_val_non_webkit, // for non webkit browsers
		'-moz-transform: scale(' + zoom_val_non_webkit + ', ' + zoom_val_non_webkit + ')',
		'-moz-transform-origin: top;',
		'-webkit-transform: (' + zoom_val_non_webkit + ', ' + zoom_val_non_webkit + ')',
		'-ms-transform: (' + zoom_val_non_webkit + ', ' + zoom_val_non_webkit + ')',
		'-o-transform: (' + zoom_val_non_webkit + ', ' + zoom_val_non_webkit + ')'
	].join(';');
	// check if current browser is IE - if browser is IE we need to add zoom styles to each document page instead of gd-panzoom element
	var ua = window.navigator.userAgent;
	var is_ie = /MSIE|Trident/.test(ua);
	if (is_ie) {
		$.each($('#gd-panzoom').find(".gd-page"), function (index, page) {
			$(page).attr('style', style);
		});
	} else {
		$('#gd-panzoom').attr('style', style);
	}
}

/**
* Get currently set zoom value
*/
function getZoomValue() {
	return parseInt($('#gd-zoom-value').text().slice(0, -1));
}

/**
* Get total matches count from search
*/
function getTotalSearchMatches() {
	return $('.gd-highlight').length;
}

/**
* Set number of currently selected match
* @param {int} index - current searched result position
* @param {int} totalCound - total search matches
*/
function setSearchMatchCount(index, totalCount) {
	$('#gd-search-count').text(index + " of " + totalCount);
}

/**
* Set zoom values on navigation panel
* @param {int} value - zoom value from 0 to 100
*/
function setNavigationZoomValues(value) {
	$('#gd-zoom-value').text(value);
}

/**
* Set page values on navigation panel
* @param {int} firstPageNumber - first page number
* @param {int} lastPageNumber - last page number or total pages
*/
function setNavigationPageValues(firstPageNumber, lastPageNumber) {
	$('#gd-page-num').text(firstPageNumber + '/' + lastPageNumber);
}

/**
* Clear previously highlighted search
*/
function clearHighlightSearch() {
	//remove highlights
	$('#gd-pages .gd-highlight').contents().unwrap();
	// normalize text
	$('#gd-pages').each(function (index, element) {
		element.normalize();
	});
	search_position = 0;
}

/**
* Clear all data from peviously loaded document
*/
function clearPageContents() {
	// set zoom to default
	setZoomValue(100);
	// set page number and total pages to zero
	setNavigationPageValues('0', '0');
	// remove previously rendered document pages
	$('#gd-panzoom').html('');
	$('#gd-thumbnails-panzoom').html('');
	// go to top
	$('#gd-pages').scrollTo(0, {
		duration: 0
	});
}

/**
* Clear all data from peviously loaded document
* @param {string} message - message to diplay in popup
*/
function printMessage(message) {
	/*var content = '<div id="gd-modal-error">' + message + '</div>';
	toggleModalDialog(true, 'Error', content);*/
	$('#alert > p')[0].innerText = message;
	$('#alert').show();
}

/**
* Scroll to page
* @param {int} pageNumber - page number where to scroll
*/
function scrollToPage(pageNumber, fast, callback) {
	// get zoom value
	var zoomValue = $('#gd-panzoom').css('zoom');
	if (typeof zoomValue === 'undefined') {
		zoomValue = 100;
	} else {
		if (parseFloat(zoomValue) != 1) {
			zoomValue = $("#gd-page-1").css("zoom") * 100 * parseFloat(zoomValue);
		} else {
			zoomValue = $("#gd-page-1").css("zoom") * 100;
		}
	}
	// scroll
	scrollOff = true;
	var scrollParams = {
		zoom: zoomValue
	};
	if (fast) {
		scrollParams.duration = 0;
	}
	$('#gd-pages').scrollTo('#gd-page-' + pageNumber, scrollParams, function() {
		scrollOff = false;
		if (typeof callback === 'function') {
			callback(scrollParams);
		}
	});
}

/**
* Rotate document pages
* @param {int} angle - document page rotation angle
*/
function rotatePages(angle) {
	// Get current page number
	var pagesAttr = $('#gd-page-num').text().split('/');
	var currentPageNumber = parseInt(pagesAttr[0]);
	// Prepare pages numbers array
	var pages = [];
	pages[0] = currentPageNumber;
	if (saveRotateState) {
		// Prepare ajax data
		var data = { guid: documentGuid, angle: angle, pages: pages, password: password };
		$.ajax({
			type: 'POST',
			url: getApplicationPath('rotateDocumentPages'),
			data: JSON.stringify(data),
			contentType: "application/json",
			success: function (returnedData) {
				if (returnedData.message != undefined) {
					// open error popup
					printMessage(returnedData.message);
					return;
				}
				$.each(returnedData, function (index, elem) {
					// Rotate the page
					rotatePage(elem.pageNumber, elem.angle);
					// set correct styles when page has landscape orientation
					if (elem.angle == 90 || elem.angle == 270) {
						if (htmlMode) {
							if ($('#gd-page-' + elem.pageNumber).width() > $('#gd-page-' + elem.pageNumber).height()) {
								$('#gd-page-' + elem.pageNumber).addClass("gd-landscape-rotated");
								$('#gd-thumbnails-page-' + elem.pageNumber).addClass("gd-thumbnails-landscape-rotated");
							} else {
								$('#gd-page-' + elem.pageNumber).addClass("gd-landscape");
								$('#gd-thumbnails-page-' + elem.pageNumber).addClass("gd-thumbnails-landscape");
							}
						} else {
							if ($('#gd-page-' + elem.pageNumber).width() > $('#gd-page-' + elem.pageNumber).height()) {
								$('#gd-page-' + elem.pageNumber).addClass("gd-landscape-image-rotated");
								$('#gd-thumbnails-page-' + elem.pageNumber).addClass("gd-thumbnails-landscape-image-rotated");
							} else {
								$('#gd-page-' + elem.pageNumber).addClass("gd-landscape-image");
								$('#gd-thumbnails-page-' + elem.pageNumber).addClass("gd-thumbnails-landscape-image");
							}
							$('#gd-thumbnails-page-' + elem.pageNumber).find("img").removeClass("gd-page-image");
						}
					} else {
						$('#gd-page-' + elem.pageNumber).removeClass("gd-landscape");
						$('#gd-thumbnails-page-' + elem.pageNumber).removeClass("gd-thumbnails-landscape");
						$('#gd-page-' + elem.pageNumber).removeClass("gd-landscape-image");
						$('#gd-thumbnails-page-' + elem.pageNumber).removeClass("gd-thumbnails-landscape-image");
					}
					// rotate page thumbnail
					rotateThumbnail(currentPageNumber, elem.angle);
				});
			},
			error: function (xhr, status, error) {
				var err = eval("(" + xhr.responseText + ")");
				console.log(err.Message);
				// open error popup
				printMessage(err.message);
			}
		});
	} else {
		documentData.pages[currentPageNumber - 1].angle = parseInt(documentData.pages[currentPageNumber - 1].angle) + parseInt(angle);
		if (documentData.pages[currentPageNumber - 1].angle > 360) {
			documentData.pages[currentPageNumber - 1].angle = 90;
		} else if (documentData.pages[currentPageNumber - 1].angle < -360) {
			documentData.pages[currentPageNumber - 1].angle = -90;
		}
		rotatePage(currentPageNumber, documentData.pages[currentPageNumber - 1].angle);
		rotateThumbnail(currentPageNumber, documentData.pages[currentPageNumber - 1].angle);
	}
}

function rotatePage(currentPageNumber, angle) {
	$('#gd-page-' + currentPageNumber).css('animation', 'none');
	$('#gd-page-' + currentPageNumber).css('transition-property', 'none');
	$('#gd-page-' + currentPageNumber).css('transform', 'rotate(' + angle + 'deg)');
}

function rotateThumbnail(currentPageNumber, angle) {
	$('#gd-thumbnails-page-' + currentPageNumber).css('animation', 'none');
	$('#gd-thumbnails-page-' + currentPageNumber).css('transition-property', 'none');
	$('#gd-thumbnails-page-' + currentPageNumber).css('transform', 'rotate(' + angle + 'deg)');
}

/**
* Download current document
*/
function downloadDocument() {
	if (documentGuid != "" && typeof documentGuid != "undefined") {
		// Open download dialog
		window.location.assign(getApplicationPath('downloadDocument/?path=') + documentGuid);
	} else {
		// open error popup
		printMessage("Please open document first");
	}
}

/**
* Upload document
* @param {file} file - File for uploading
* @param {string} url - URL of the file, set it if URL used instead of file
*/
function uploadDocument(file, url) {
	// prepare form data for uploading
	var formData = new FormData();
	// add local file for uploading
	formData.append("file", file);
	// add URL if set
	if (typeof url != "undefined" && url != null) {
		formData.append("url", url);
	}
	formData.append("rewrite", rewrite);
	$.ajax({
		// callback function which updates upload progress bar
		xhr: function () {
			var xhr = new window.XMLHttpRequest();
			// upload progress
			xhr.upload.addEventListener("progress", function (event) {
				if (event.lengthComputable) {
					$(".gd-modal-close-action").off('click');
					$("#gd-open-document").prop("disabled", true);
					if (event.loaded == event.total) {
						$('.gd-modal-close-action').on('click', closeModal);
						$("#gd-open-document").prop("disabled", false);
						if ($(".gd-checkbox").length > 0) {
							loadFileTree("", true);
						} else {
							loadFileTree("");
						}
					}
				}
			}, false);
			return xhr;
		},
		type: 'POST',
		url: getApplicationPath('uploadDocument'),
		data: formData,
		cache: false,
		contentType: false,
		processData: false,
		success: function (returnedData) {
			if (returnedData.message != undefined) {
				// open error popup
				printMessage(returnedData.message);
				return;
			}
		},
		error: function (xhr, status, error) {
			var err = eval("(" + xhr.responseText + ")");
			console.log(err.Message);
			// open error popup
			printMessage(err.message);
		}
	});
}

/**
* Print current document
*/
function printDocument(event) {
	event.stopImmediatePropagation();
	if ($(this).find("li").length > 0) {
		return;
	}
	loadPrint();
	gaEvent('Print');
}

/**
* Close modal
*/
function closeModal() {
	$("#gd-upload-input").val('');
	toggleModalDialog(false, '');
}

function openBrowseModal() {
	loadFileTree('');
}

/**
* Open password modal
* @param {string} error - error message
**/
function openPasswordModal(error) {
	var passwordSection = '<section id="gd-password-section" class="tab-slider-body">' +
		'<div class="inner-addon left-addon btn gd-password-wrap" id="gd-password-wrap">' +
		'<input type="password" class="form-control" id="gd-password-input" placeholder="Enter password">' +
		'<button class="btn btn-primary" id="gd-password-submit">Submit</button>' +
		'<span class="gd-password-error" style="display: none;"></span>' +
		'</div>' +
		'</section>';
	toggleModalDialog(true, 'Password required', passwordSection);
	if (error != "" && typeof error != "undefined") {
		$(".gd-password-error")[0].innerHTML = error;
		$(".gd-password-error").show();
	} else {
		$(".gd-password-error").hide();
	}
}

/**
* Check if all document pages are loaded and clear the password
**/
function clearPassword() {
	if (password != "" && typeof password != "undefined") {
		var totalPagesNumber = $('#gd-panzoom > div').length;
		var loadedPages = [];
		if (totalPagesNumber > 0) {
			$('#gd-panzoom > div').each(function (index, page) {
				if ($(page).has(".gd-wrapper").length) {
					loadedPages.push(true);
				}
			});
			if (loadedPages.length == totalPagesNumber) {
				password = "";
			}
		}
	}
}

/**
* On-promise function which waits until the element is loaded
* @param {Object} selector - element to wait for
**/
function isPageLoaded(selector) {
	return new Promise(function (resolve, reject) {
		// check if loaded       
		var waitForEl = function (selector, count) {
			var count = 0;
			var el = selector.find(".gd-wrapper");
			// check if element is loaded
			if (el.length > 0) {
				resolve(el);
			} else {
				// wait 100 milliseconds and check again
				setTimeout(function () {
					if (typeof count != "undefined" && count != null) {
						count++;
						if (count < 120) {
							waitForEl(selector, count);
						} else {
							reject();
						}
					}
				}, 1000);
			}
		};
		waitForEl(selector);
	});
}

/**
* Set zoom level
*/
function setZoomLevel(zoomString) {
	var zoomValue = zoomString || $(this).text();
	var fitHeight = false;
	switch (zoomValue) {
		case 'Fit Width':
			// get page width
			var pageWidth = $('#gd-pages .gd-page').width() + 30;
			// get screen width
			var screenWidth = $('#gd-pages').width();
			// get scale ratio
			var scale = (pageWidth / screenWidth) * 100;
			// set values
			zoomValue = (screenWidth / pageWidth) * 100;
			break;
		case 'Fit Height':
			fitHeight = true;
			// get page height
			var pageHeight = $('#gd-pages .gd-page').height() + 30;
			// get screen height
			var screenHeight = $('#gd-pages').height();
			// get scale ratio
			var scale = (screenHeight / pageHeight) * 100;
			// set values
			zoomValue = scale;
			break;
		default:
			zoomValue = zoomValue.slice(0, -1);
			break;
	}
	setZoomValue(zoomValue);

	if (fitHeight) {
		var pagesAttr = $('#gd-page-num').text().split('/');
		// $('#gd-pages').scrollTo('#gd-page-' + pagesAttr[0], {duration: 0});
		window.location.href = '#gd-page-' + pagesAttr[0];
	}

	gaEvent('Zoom Value');
}

/**
* Zoom in
*/
function zoomIn(event) {
	event.stopImmediatePropagation();
	var zoom_val = getZoomValue();
	if (zoom_val < 490) {
		zoom_val = zoom_val + 10;
	}
	setZoomValue(zoom_val);
	gaEvent('Zoom In');
}

/**
* Zoom out
*/
function zoomOut(event) {
	event.stopImmediatePropagation();
	var zoom_val = getZoomValue();
	if (zoom_val > 30) {
		zoom_val = zoom_val - 10;
	}
	setZoomValue(zoom_val);
	gaEvent('Zoom Out');
}

/**
* Switch to thumbnails
*/
function panelSwitchToThumbnails() {
	$('#gd-thumbnails-tab-1').addClass('active');
	$('#gd-thumbnails-tab-2').removeClass('active');

	$('#gd-thumbnails-panzoom').removeClass('hidden');
	$('#gd-thumbnails-headers').addClass('hidden');
}

/**
* Switch to headers
*/
function panelSwitchToHeaders() {
	$('#gd-thumbnails-tab-1').removeClass('active');
	$('#gd-thumbnails-tab-2').addClass('active');

	$('#gd-thumbnails-panzoom').addClass('hidden');
	$('#gd-thumbnails-headers').removeClass('hidden');
}

/**
* Get HTML content for file browser modal
**/
function getHtmlFileBrowser(multiple) {
	var multipleActions = "";
	var single = "single";
	if (multiple) {
		multipleActions = '<input type="checkbox" class="gd-select-all gd-checkbox">' +
			'<div class="gd-add-selected"><i class="fa fa-plus"></i><label>Add selected</label></div>';
		single = "";
	}
	var uploadButtons = "";
	if (upload) {
		uploadButtons = '<label class="gd-upload-dropdown">' +
			'<div class="gd-button">' +
			'<i class="fa fa-upload"></i><label>Upload file</label><i class="fas fa-chevron-down"></i>' +
			'</div>' +
			'<input type="checkbox" class="gd-upload-input" id="gd-upload-input-checkbox">' +
			'<ul class="gd-upload-menu">' +
			'<li>Upload from:</li>' +
			'<li><label><i class="fas fa-hdd"></i> Disc <input id="gd-upload-input" type="file" multiple style="display: none;" ></label></li>' +
			'<li id="gd-upload-url"><i class="fas fa-link"></i> URL</li>' +
			'</ul>' +
			'</label>' +
			'<div class="inner-addon left-addon btn gd-url-wrap" id="gd-url-wrap" style="">' +
			'<input type="url" class="form-control" id="gd-url" placeholder="http://">' +
			'<button class="btn" id="gd-add-url"><i class="fas fa-check"></i></button>' +
			'</div>' +
			'<i class="fas fa-hdd gd-hdd"></i>' +
			'<i class="fab fa-aws gd-aws"></i>';
	}
	var uploadButtons = '<div id="gd-upload-actions" class="gd-action-buttons ' + single + '">' +
		multipleActions +
		uploadButtons +
		'</div>';

	return uploadButtons + '<section id="gd-browse-section" class="tab-slider-body">' +
		'<div id="gd-modal-spinner"><i class="fa fa-circle-o-notch fa-spin"></i> &nbsp;Loading... Please wait.</div>' +
		'<div id="gd-modal-filebroswer" class="gd-modal-table">' +
		'<div class="gd-modal-table-head ' + single + '"><label>Document</label><label class="' + single + '">Size</label></div>' +
		'<div class="gd-modal-table-body">' +
		'<div class="text-center gd-go-up ' + single + '"><i class="fa fa-level-up"></i></div>' +
		'<div class="gd-filetree-up gd-go-up ' + single + '">...</div>' +
		'</div>' +
		'</div>' +
		'</section>' +
		'<div class="gd-drag-n-drop-wrap" id="gd-dropZone">' +
		'<i class="fa fa-cloud-download fa-5x" aria-hidden="true"></i>' +
		'<h2>Drag &amp; Drop your files here</h2>' +
		'</div>';
}

/**
* INit drag N Drop zone
* @param {Object} uploadAction - function used to upload droped file
**/
function initDragNDrop(uploadAction) {
	var dropZone = $('#gd-dropZone');
	if (typeof dropZone[0] != "undefined") {
		//Drag n drop functional
		if ($('#gd-dropZone').length) {
			if (typeof (window.FileReader) == 'undefined') {
				dropZone.text("Your browser doesn't support Drag and Drop");
				dropZone.addClass('error');
			}
		}

		$(".gd-modal-body").on("dragover", function () {
			dropZone.show();
			return false;
		});

		dropZone[0].ondragleave = function () {
			dropZone.hide();
			return false;
		};

		dropZone[0].ondrop = function (event) {
			event.preventDefault();
			var files = event.dataTransfer.files;
			$.each(files, function (index, file) {
				uploadAction(file);
			});
			dropZone.hide();
		};
	}
}

/*
******************************************************************
******************************************************************
GROUPDOCS.VIEWER PLUGIN
******************************************************************
******************************************************************
*/
(function ($) {
	/*
	******************************************************************
	STATIC VALUES
	******************************************************************
	*/
	var gd_navbar = '#gd-navbar';

	/*
	******************************************************************
	METHODS
	******************************************************************
	*/
	var methods = {
		init: function (options) {
			// set defaults
			var defaults = {
				applicationPath: null,
				preloadPageCount: 1,
				zoom: true,
				pageSelector: true,
				search: false,
				thumbnails: true,
				rotate: true,
				download: true,
				upload: true,
				print: true,
				defaultDocument: null,
				browse: true,
				rewrite: true,
				htmlMode: true,
				saveRotateState: true,
				enableRightClick: true
			};
			options = $.extend(defaults, options);

			// set global option params
			applicationPath = options.applicationPath;
			preloadPageCount = options.preloadPageCount;
			rewrite = options.rewrite;
			htmlMode = options.htmlMode;
			thumbnails = options.thumbnails;
			saveRotateState = options.saveRotateState;
			enableRightClick = options.enableRightClick;
			upload = options.upload;
			// assembly html base
			this.append(getHtmlBase);
			this.append(getHtmlModalDialog);

			// assembly nav bar
			if (options.browse) {
				$(gd_navbar).append(getHtmlBrowsePanel);
				$("#gd-btn-browse").on('click', openBrowseModal);
			}
			if (options.zoom) {
				$(gd_navbar).append(getHtmlNavZoomPanel);
				$(gd_navbar).append(getHtmlNavSplitter);
			}
			if (options.pageSelector) {
				$(gd_navbar).append(getHtmlNavPagesPanel);
				$(gd_navbar).append(getHtmlNavSplitter);
			}
			if (options.rotate) {
				$(gd_navbar).append(getHtmlRotatePanel);
				$(gd_navbar).append(getHtmlNavSplitter);
			}
			if (options.search) {
				$(gd_navbar).append(getHtmlNavSearchPanel);
				$(gd_navbar).append(getHtmlNavSplitter);
			}
			if (options.download) {
				$(gd_navbar).append(getHtmlNavDownloadPanel);
				$(gd_navbar).append(getHtmlNavSplitter);
			}
			if (options.print) {
				$(gd_navbar).append(getHtmlNavPrintPanel);
				$(gd_navbar).append(getHtmlNavSplitter);
			}
			if (options.thumbnails) {
				$(gd_navbar).append(getHtmlNavThumbTogglePanel);
			}
			if (options.defaultDocument) {
				documentGuid = options.defaultDocument;
				loadDocument(function (data) {
					generatePagesTemplate(data);
				});
			}
		}
	};


	/*
	******************************************************************
	INIT PLUGIN
	******************************************************************
	*/
	$.fn.viewer = function (method) {
		if (methods[method]) {
			return methods[method].apply(this, Array.prototype.slice.call(arguments, 1));
		} else if (typeof method === 'object' || !method) {
			return methods.init.apply(this, arguments);
		} else {
			$.error('Method' + method + ' does not exist on jQuery.viewer');
		}
	};


	/*
	******************************************************************
	HTML MARKUP
	******************************************************************
	*/
	function getHtmlBase() {
		return '<div id="gd-container">' +
			'<div class="wrapper">' +
			// header BEGIN
			'<div id="gd-header">' +
			//'<div id="gd-header-logo">' +
			//'</div>' +

			// nav bar BEGIN
			'<ul id="' + gd_navbar.slice(1) + '">' +
			// nav bar content
			'</ul>' +
			// nav bar END
			'</div>' +
			// header END

			// thumbnails sidebar BEGIN
			'<div id="gd-thumbnails">' +
			'<div class="gd-thumbnails-toggler" onclick="toggleThumbnails()" title="Toggle sidebar"></div>' +
			'<h3>Navigation</h3>' +
			'<ul class="tabs">' +
			'<li id="gd-thumbnails-tab-2" class="active" onclick="panelSwitchToHeaders()"><a href="javascript:void(0)">Headers</a></li>' +
			'<li id="gd-thumbnails-tab-1" onclick="panelSwitchToThumbnails()"><a href="javascript:void(0)">Thumbnails</a></li>' +
			'</ul>' +
			'<div id="gd-thumbnails-headers">'+
			// Headers will be added here automatically on document open
			'</div>' +
			'<div id="gd-thumbnails-panzoom" class="hidden">' +
			// Thumbnails will be added here automatically on document open
			'</div>' +
			'</div>' +
			// thumbnails sidebar END

			// pages BEGIN
			'<div id="gd-pages">' +
			'<div id="gd-panzoom">' +
			// list of pages
			'</div>' +
			'</div>' +
			// pages END

			'</div>' +
			'</div>' +
			'<div id="gd-container-fade" class="gd-container-fade">' +
			//'<div id="gd-container-fade-text" class="gd-container-spinner"><i class="fa fa-circle-o-notch fa-spin"></i> &nbsp;Loading... Please wait.</div>' +
			'</div>' +
			'<div id="preloader" class="hidden"><div class="spinner"></div></div>';
	}

	function getHtmlModalDialog() {
		return '<div class="gd-modal fade" id="modalDialog">' +
			'<div class="gd-modal-dialog">' +
			'<div class="gd-modal-content" id="gd-modal-content">' +
			// header
			'<div class="gd-modal-header">' +
			'<div class="gd-modal-close gd-modal-close-action"><span>&times;</span></div>' +
			'<h4 class="gd-modal-title"></h4>' +
			'</div>' +
			// body
			'<div class="gd-modal-body">' +
			// modal content will be here
			'</div>' +
			// footer
			'<div class="gd-modal-footer">' +
			// empty footer
			'</div>' +
			'</div><!-- /.modal-content -->' +
			'</div><!-- /.modal-dialog -->' +
			'</div>';
	}

	function getHtmlNavSplitter() {
		return '<li class="gd-nav-separator" role="separator"></li>';
	}

	function getHtmlNavZoomPanel() {
		return '<li class="gd-nav-toggle" id="gd-zoom-val-container">' +
			'<span id="gd-zoom-value">100%</span>' +
			'<span class="gd-nav-caret"></span>' +
			'<ul class="gd-nav-dropdown-menu gd-nav-dropdown disabled" id="gd-btn-zoom-value">' +
			'<li>25%</li>' +
			'<li>50%</li>' +
			'<li>100%</li>' +
			'<li>150%</li>' +
			'<li>200%</li>' +
			'<li>300%</li>' +
			'<li role="separator" class="gd-nav-dropdown-menu-separator"></li>' +
			'<li>Fit Width</li>' +
			'<li>Fit Height</li>' +
			'</ul>' +
			'</li>' +
			'<li id="gd-btn-zoom-in" class="disabled">' +
			'<i class="fa fa-search-plus"></i>' +
			'<span class="gd-tooltip">Zoom In</span>' +
			'</li>' +
			'<li id="gd-btn-zoom-out" class="disabled">' +
			'<i class="fa fa-search-minus"></i>' +
			'<span class="gd-tooltip">Zoom Out</span>' +
			'</li>';
	}

	function getHtmlNavPagesPanel() {
		return '<li id="gd-btn-page-first" class="gd-nav-btn-pages">' +
			'<i class="fa fa-angle-double-left"></i>' +
			'<span class="gd-tooltip">First Page</span>' +
			'</li>' +
			'<li id="gd-btn-page-prev" class="gd-nav-btn-pages">' +
			'<i class="fa fa-angle-left"></i>' +
			'<span class="gd-tooltip">Previous Page</span>' +
			'</li>' +
			'<li id="gd-page-num">0/0</li>' +
			'<li id="gd-btn-page-next" class="gd-nav-btn-pages">' +
			'<i class="fa fa-angle-right"></i>' +
			'<span class="gd-tooltip">Next Page</span>' +
			'</li>' +
			'<li id="gd-btn-page-last" class="gd-nav-btn-pages">' +
			'<i class="fa fa-angle-double-right"></i>' +
			'<span class="gd-tooltip">Last Page</span>' +
			'</li>';
	}

	function getHtmlNavSearchPanel() {
		return '<li id="gd-btn-search" class="gd-nav-toggle">' +
			'<i class="fa fa-search"></i>' +
			'<span class="gd-tooltip">Search</span>' +
			'<div id="gd-nav-search-container" class="gd-nav-dropdown">' +
			'<input type="text" id="gd-search-input"/>' +
			'<div id="gd-search-count">0 of 0</div>' +
			'<div class="gd-nav-search-btn" id="gd-nav-search-prev"><i class="fa fa-chevron-left"></i></div>' +
			'<div class="gd-nav-search-btn" id="gd-nav-search-next"><i class="fa fa-chevron-right"></i></div>' +
			'<div class="gd-nav-search-btn" id="gd-nav-search-cancel"><i class="fa fa-times"></i></div>' +
			'</div>' +
			'</li>';
	}

	function getHtmlNavThumbTogglePanel() {
		return '<li id="gd-nav-right"><i class="fa fa-th-large"></i><span class="gd-tooltip">Navigation</span></li>';
	}

	function getHtmlRotatePanel() {
		return '<li id="gd-btn-counterclockwise"><i class="fa fa-rotate-left"></i><span class="gd-tooltip">Rotate CCW</span></li>' +
			'<li id="gd-btn-clockwise"><i class="fa fa-rotate-right"></i><span class="gd-tooltip">Rotate CW</span></li>';
	}

	function getHtmlNavDownloadPanel() {
		return '<li id="gd-btn-download" class="disabled"><i class="fa fa-download"></i><span class="gd-tooltip">Download</span></li>';
	}

	function getHtmlNavPrintPanel() {
		return '<li id="gd-btn-print" class="disabled"><i class="fa fa-print"></i><span class="gd-tooltip">Print</span></li>';
	}

	function getHtmlBrowsePanel() {
		return '<li id="gd-btn-browse"><i class="fa fa-folder-open"></i><span class="gd-tooltip">Browse files</span></li>';
	}
})(jQuery);

/*
******************************************************************
******************************************************************
JQUERY SCROLL TO PLUGIN
******************************************************************
******************************************************************
*/
$.fn.scrollTo = function (target, options, callback) {
	if (typeof options == 'function' && arguments.length == 2) { callback = options; options = target; }
	var settings = $.extend({
		scrollTarget: target,
		offsetTop: 100,
		duration: 500,
		zoom: options.zoom,
		easing: 'swing'
	}, options);
	return this.each(function () {
		var scrollPane = $(this);
		if (isNaN(settings.zoom)) {
			settings.zoom = 100;
		}
		var scrollTarget = (typeof settings.scrollTarget == "number") ? settings.scrollTarget : $(settings.scrollTarget);
		if (typeof settings.scrollTarget != "number") {
			var scrollYTop = scrollTarget.offset().top * (settings.zoom / 100);
		}
		var scrollY = (typeof scrollTarget == "number") ? scrollTarget : scrollYTop + scrollPane.scrollTop() - parseInt(settings.offsetTop);
		scrollPane.animate({ scrollTop: scrollY }, parseInt(settings.duration), settings.easing, function () {
			if (typeof callback == 'function') {
				callback.call(this);
			}
		});
	});
};

/*
******************************************************************
******************************************************************
JQUERY CHECK IF IN VIEWPORT PLUGIN
******************************************************************
******************************************************************
*/
$.fn.isOnScreen = function (x, y) {

	if (x == null || typeof x == 'undefined') x = 1;
	if (y == null || typeof y == 'undefined') y = 1;

	var win = $(window);

	var viewport = {
		top: win.scrollTop(),
		left: win.scrollLeft()
	};
	viewport.right = viewport.left + win.width();
	viewport.bottom = viewport.top + win.height();
	var zoom = parseInt($("#gd-zoom-value").html()) / 100;
	if (isNaN(zoom)) {
		zoom = 1;
	}
	var height = this.outerHeight() * zoom;
	var width = this.outerWidth() * zoom;

	if (!width || !height) {
		return false;
	}

	var bounds = this.offset();
	bounds.right = (bounds.left * zoom) + width;
	bounds.bottom = (bounds.top * zoom) + height;

	var visible = (!(viewport.right < (bounds.left * zoom) || viewport.left > bounds.right || viewport.bottom < (bounds.top * zoom) || viewport.top > bounds.bottom));

	if (!visible) {
		return false;
	}

	var deltas = {
		top: Math.min(1, (bounds.bottom - viewport.top) / height),
		bottom: Math.min(1, (viewport.bottom - (bounds.top * zoom)) / height),
		left: Math.min(1, (bounds.right - viewport.left) / width),
		right: Math.min(1, (viewport.right - (bounds.left * zoom)) / width)
	};

	return (deltas.left * deltas.right) >= x && (deltas.top * deltas.bottom) >= y;
};

/*
******************************************************************
******************************************************************
CHECK IF MOBILE
******************************************************************
******************************************************************
*/
var isMobile = function () {
	return navigator.maxTouchPoints > 0 || //for chrome
		window.navigator.msMaxTouchPoints > 0 ||
		'ontouchstart' in window; // works on most browsers
};

/*
******************************************************************
******************************************************************
SEND GOOGLE ANALYTICS EVENT
******************************************************************
******************************************************************
*/

var gaEvent = function(action, category) {
	if (!category) {
		category = 'Viewer Navbar';
	}
	var ga = ('ga' in window.parent) ? window.parent.ga : (('ga' in window) ? window.ga : null);
	if (ga !== null) {
		try {
			var tracker = ga.getAll()[0];
			if (tracker !== undefined) {
				tracker.send('event', {
					'eventCategory' : category,
					'eventAction'   : action
				});
			}
		} catch (e) {}
	}
}