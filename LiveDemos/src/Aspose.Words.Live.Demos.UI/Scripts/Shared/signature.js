var drawCanvas = null;
var pickerDraw = null;
var pickerText = null;
var points_count = 0;

function setSettings(kind) {
	switch (kind) {
		case "drawing":
			$("#drawingSettings").show();
			$("#textSettings").hide();
			$("#imageSettings").hide();
			break;
		case "text":
			$("#drawingSettings").hide();
			$("#textSettings").show();
			$("#imageSettings").hide();
			break;
		case "image":
			$("#drawingSettings").hide();
			$("#textSettings").hide();
			$("#imageSettings").show();
			break;
	}
}

function AssignSaveToButtonText(obj) {
	var t = $(obj).text();
	document.getElementById('SaveToButton').innerHTML = t;
	document.getElementById('SaveToHidden').value = t;
}

function setSignatureType(obj) {
	var t = $(obj).text();
	document.getElementById('signatureType').innerHTML = t;
	o.signatureType = $(obj).data('val');
	setSettings(o.signatureType);
}

function removefileimage() {
	$('#fileUploadImage').hide();
	$('#fileUploadImageInput').show();
}

function Canvas(signCanvas, pickcolor) {
	var signatureCanvas = document.querySelector(signCanvas);
	if (signatureCanvas === null)
		return;
	var ctx = signatureCanvas.getContext('2d');

	var mouse = { x: null, y: null };
	var last_mouse = { x: 0, y: 0 };

	function resetCanvasContext(color) {
		ctx.lineWidth = 2;
		ctx.lineJoin = 'round';
		ctx.lineCap = 'round';
		ctx.strokeStyle = color;
	}

	resetCanvasContext($(pickcolor).val());

	$(pickcolor).change(function () {
		points_count = 0;
		// clearSignature();
		resetCanvasContext($(this).val());
	}
	);

	// set the listener on pointer down event
	var setOnDownListener = function (handlerDown, handlerMove) {
		signatureCanvas.addEventListener(handlerDown, function (ev) {
			signatureCanvas.addEventListener(handlerMove, onPaint, false);
		}, false);
	};

	// set the listener on pointer up event
	var setOnUpListener = function (handlerUp, handlerMove) {
		signatureCanvas.addEventListener(handlerUp, function () {
			mouse.x = null;
			mouse.y = null;
			signatureCanvas.removeEventListener(handlerMove, onPaint, false);
		}, false);
	};

	// set the listener on pointer move event
	var setOnMoveListener = function (handlerMove) {
		signatureCanvas.addEventListener(handlerMove, function (ev) {
			last_mouse.x = mouse.x;
			last_mouse.y = mouse.y;

			var bRect = signatureCanvas.getBoundingClientRect();
			mouse.x = (ev.clientX - bRect.left) * (signatureCanvas.width / bRect.width);
			mouse.y = (ev.clientY - bRect.top) * (signatureCanvas.height / bRect.height);
		}, false);
	};

	// check if the specified event handler is supported
	var isEventSupported = function (eventName) {
		var cnv = document.createElement('canvas');
		var isSupported = (eventName in cnv);
		if (!isSupported) {
			cnv.setAttribute(eventName, 'return;');
			isSupported = typeof cnv[eventName] === 'function';
		}
		cnv = null;
		return isSupported;
	};

	// handler names on pointer events
	var handlerDown = null, handlerUp = null, handlerMove = null;

	// add event listeners
	if (isEventSupported('onpointerdown')) {
		handlerDown = 'pointerdown';
		handlerUp = 'pointerup';
		handlerMove = 'pointermove';
	} else if (isEventSupported('ontouchstart')) {
		handlerDown = 'touchstart';
		handlerUp = 'touchend';
		handlerMove = 'touchmove';
	} else if (isEventSupported('onmousedown')) {
		handlerDown = 'mousedown';
		handlerUp = 'mouseup';
		handlerMove = 'mousemove';
	}
	if (handlerDown !== null && handlerUp !== null && handlerMove !== null) {
		setOnDownListener(handlerDown, handlerMove);
		setOnUpListener(handlerUp, handlerMove);
		setOnMoveListener(handlerMove);
	}

	var onPaint = function () {
		if (last_mouse.x === null && last_mouse.y === null) {
			last_mouse.x = mouse.x;
			last_mouse.y = mouse.y;
		}
		ctx.beginPath();
		ctx.moveTo(last_mouse.x, last_mouse.y);
		points_count++;
		ctx.lineTo(mouse.x, mouse.y);
		ctx.closePath();
		ctx.stroke();
	};

	this.clearSignature = function () {
		points_count = 0;
		signatureCanvas.width = signatureCanvas.width;
		var ctx = signatureCanvas.getContext('2d');
		ctx.lineWidth = 2;
		ctx.lineJoin = 'round';
		ctx.lineCap = 'round';
		ctx.strokeStyle = $(pickcolor).val();
		mouse.x = null;
		mouse.y = null;
	};
}

function colorPicker(color_picker) {
	var colorList = [
		'000000', '993300', '333300', '003300', '003366', '000066', '333399', '333333',
		'660000', 'FF6633', '666633', '336633', '336666', '0066FF', '666699', '666666', 'CC3333', 'FF9933', '99CC33',
		'669966', '66CCCC', '3366FF', '663366', '999999', 'CC66FF', 'FFCC33', 'FFFF66', '99FF66', '99CCCC', '66CCFF',
		'993366', 'CCCCCC', 'FF99CC', 'FFCC99', 'FFFF99', 'CCffCC', 'CCFFff', '99CCFF', 'CC99FF', 'FFFFFF'
	];

	var picker = $(color_picker);

	for (var i = 0; i < colorList.length; i++) {
		picker.append('<li class="color-item" data-hex="' + '#' + colorList[i] + '" style="background-color:' + '#' + colorList[i] + ';"></li>');
	}

	$('body').click(function () {
		picker.fadeOut();
	});

	picker.siblings('.call-picker').click(function (event) {
		event.stopPropagation();
		picker.fadeIn();
		picker.children('li').hover(function () {
			var codeHex = $(this).data('hex');

			picker.siblings('.color-holder').css('background-color', codeHex);
			picker.siblings('input.call-picker').val(codeHex).trigger('change');
		});
	});
	return picker;
}

function validateSignature() {
	switch (o.signatureType) {
		case "drawing":
			if (points_count === 0) {
				showAlert(o.validationDrawMessage);
				return false;
			}
			break;
		case "text":
			var t = $('#signText').val();
			if (t.length === 0) {
				showAlert(o.validationTextMessage);
				return false;
			}
			break;
		case "image":
			var files = $('#fileUploadImageInput')[0].files;
			if (files === undefined || files.length === 0) {
				showAlert(o.FileSelectMessage);
				return false;
			}
			break;
	}
	return true;
}


function requestSignature() {
	if (!validateSignature())
		return;

	let data = fileDrop.prepareFormData();
	if (data === null)
		return;

	switch (o.signatureType) {
		case "drawing":
			var image = $('#signCanvas')[0].toDataURL("image/png");
			image = image.replace('data:image/png;base64,', '');
			data.append('image', image);
			break;
		case "text":
			data.append('text', $('#signText').val());
			data.append('textColor', $('#pickcolorText').val());
			break;
		case "image":
			var file = $('#fileUploadImageInput')[0].files[0];
			data.append('imageFile', file, file.name);
			break;
	}

	let url = o.UIBasePath +
		'api/AsposeWordsSignature/Sign?outputType=' + $('#saveAs').val() +
		'&signatureType=' + o.signatureType;
	request(url, data);
}


$(document).ready(function () {
	$('#fileUploadImage').hide();

	$('#fileUploadImageInput').change(function () {
		var file = $('#fileUploadImageInput')[0].files[0].name;
		$('#fileUploadImage label').text(file);
		$('#fileUploadImage').show();
	});
	drawCanvas = new Canvas('#signCanvas', "#pickcolorDrawing");
	pickerDraw = new colorPicker('#color-pickerDraw');
	pickerText = new colorPicker('#color-pickerText');
	setSettings(o.signatureType);
});