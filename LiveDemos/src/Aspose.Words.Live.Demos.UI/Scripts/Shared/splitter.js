function validateSplitter(splitType, pars) {
	var result = true;
	switch (splitType) {
		case '3':
			if (pars.length === 0 || pars.match(/^\s+$/)) {
				$('#splitNumber').css({ 'border': '2px solid red' });
				showAlert(o.validationNumberMessage);
				result = false;
			}
			break;
		case '4':
			var checkResult = true;
			pars = pars.replace(/\s/g, ""); // remove whitespace
			if (pars.length === 0 || pars.match(/^\s+$/) || (pars.indexOf(',') === -1 && !pars.match(/^\d+$/) && !pars.match(/^\d+\-\d+$/))) {
				checkResult = false;
			} else {
				var splitResult = pars.split(',');
				for (var i = 0; i < splitResult.length; i++) {
					if (!splitResult[i].match(/^\d+$/) && !splitResult[i].match(/^\d+\-\d+$/)) {
						checkResult = false;
						break;
					} else if (splitResult[i].match(/^\d+\-\d+$/)) {
						var dashPos = splitResult[i].indexOf('-');
						var v1 = parseInt(splitResult[i].substring(0, dashPos));
						var v2 = parseInt(splitResult[i].substring(dashPos + 1));
						if (v1 >= v2) {
							checkResult = false;
							break;
						}
					}
				}
			}
			if (!checkResult) {
				$('#splitRange').css({ 'border': '2px solid red' });
				showAlert(o.validationRangeMessage);
				result = false;
			}
			break;
	}
	return result;
}

function requestSplitter() {
	let splitType = $('input[name=splitType]:checked').val();
	let pars = '';
	switch (splitType) {
		case '3':
			pars = $('#splitNumber').val();
			break;
		case '4':
			pars = $('#splitRange').val();
			break;
	}

	if (!validateSplitter(splitType, pars))
		return;

	let data = fileDrop.prepareFormData();
	if (data === null)
		return;

	let url = o.UIBasePath + 'Splitter/Splitter?outputType=' + $('#saveAs').val()
		+ '&splitType=' + splitType
		+ '&pars=' + pars;
	request(url, data);
}

$(document).ready(function () {
	$("#splitNumber").inputFilter(function (value) {
		return /^\d*$/.test(value); // Allow digits only, using a RegExp
		});
	$("#splitNumber").click(() => {
		$("#splitNumber").css({ 'border': '' });
		$('input[name=splitType][value=3]').attr('checked', true);
		hideAlert();
	});
	$("#splitRange").inputFilter(function (value) {
		return /^[\s,\,\-,\d]*$/.test(value); // Allow digits, whitespace and '-' only 
		});
	$("#splitRange").click(() => {
		$("#splitRange").css({ 'border': '' });
		$('input[name=splitType][value=4]').attr('checked', true);
		hideAlert();
	});
});