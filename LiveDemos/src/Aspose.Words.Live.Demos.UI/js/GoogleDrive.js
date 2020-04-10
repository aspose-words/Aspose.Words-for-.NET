// The Browser API key obtained from the Google API Console.
// Replace with your own Browser API key, or your own key.
var developerKey = document.getElementById("hdnGoogleDeveloperID").value;

// The Client ID obtained from the Google API Console. Replace with your own Client ID.
var clientId = document.getElementById("hdnGoogleClientID").value;

// Replace with your own project number from console.developers.google.com.
// See "Project number" under "IAM & Admin" > "Settings"
var appId = document.getElementById("hdnGoogleAppKey").value;

// Scope to use to access user's Drive items.
var scope = ['https://www.googleapis.com/auth/drive'];

var pickerApiLoaded = false;
var oauthToken;

// Use the Google API Loader script to load the google.picker script.
function loadPicker() {
	gapi.load('auth', { 'callback': onAuthApiLoad });
	gapi.load('picker', { 'callback': onPickerApiLoad });
}

function onAuthApiLoad() {
	window.gapi.auth.authorize(
		{
			'client_id': clientId,
			'scope': scope,
			'immediate': false
		},
		handleAuthResult);
}

function onPickerApiLoad() {
	pickerApiLoaded = true;
	createPicker();
}

function handleAuthResult(authResult) {
	if (authResult && !authResult.error) {
		oauthToken = authResult.access_token;
		createPicker();
	}
}

// Create and render a Picker object for searching images.
function createPicker() {
	if (pickerApiLoaded && oauthToken) {
		var view = new google.picker.View(google.picker.ViewId.DOCS);		
		view.setMimeTypes(document.getElementById("hdnSupportedMimeTypes").value);
		//view.setMimeTypes("application / pdf");
		var picker = new google.picker.PickerBuilder()
			.enableFeature(google.picker.Feature.NAV_HIDDEN)
			.enableFeature(google.picker.Feature.MULTISELECT_ENABLED)
			.setAppId(appId)
			.setOAuthToken(oauthToken)
			.addView(view)
			.addView(new google.picker.DocsUploadView())
			.setDeveloperKey(developerKey)
			.setCallback(pickerCallback)
			.build();
		picker.setVisible(true);
	}
}
// A simple callback implementation.
function pickerCallback(data) {
	if (data.action == google.picker.Action.PICKED) {
		var ext = data.docs[0].name.split('.').pop();
		document.getElementById('hdnFileName').value = data.docs[0].name;
		$('.filename label').text(data.docs[0].name);
		$('.fileupload').show();
		$('.filesendemail').hide();
		var upperext = ext.toUpperCase();	

		var elementhdnToValue = document.getElementById('hdnToValue');
		if (typeof (elementhdnToValue) != 'undefined' && elementhdnToValue != null) {
			elementhdnToValue.value = upperext;
		}

		var elementbtnTo = document.getElementById('btnTo');
		if (typeof (elementbtnTo) != 'undefined' && elementbtnTo != null) {
			elementbtnTo.innerHTML = upperext;
		}
		
		document.getElementById("hdnFileID").value = data.docs[0].id;		
		document.getElementById("hdnFileFrom").value = "google";		
	}	
}