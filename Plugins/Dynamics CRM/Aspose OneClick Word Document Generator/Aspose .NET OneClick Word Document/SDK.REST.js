// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================
// <snippetSDKRESTJS>
if (typeof (SDK) == "undefined")
{ SDK = { __namespace: true }; }
SDK.REST = {
 _context: function () {
  ///<summary>
  /// Private function to the context object.
  ///</summary>
  ///<returns>Context</returns>
  if (typeof GetGlobalContext != "undefined")
  { return GetGlobalContext(); }
  else {
   if (typeof Xrm != "undefined") {
    return Xrm.Page.context;
   }
   else
   { throw new Error("Context is not available."); }
  }
 },
 _getClientUrl: function () {
  ///<summary>
  /// Private function to return the server URL from the context
  ///</summary>
  ///<returns>String</returns>
  var clientUrl = this._context().getClientUrl()

  return clientUrl;
 },
 _ODataPath: function () {
  ///<summary>
  /// Private function to return the path to the REST endpoint.
  ///</summary>
  ///<returns>String</returns>
  return this._getClientUrl() + "/XRMServices/2011/OrganizationData.svc/";
 },
 _errorHandler: function (req) {
  ///<summary>
  /// Private function return an Error object to the errorCallback
  ///</summary>
  ///<param name="req" type="XMLHttpRequest">
  /// The XMLHttpRequest response that returned an error.
  ///</param>
  ///<returns>Error</returns>
  //Error descriptions come from http://support.microsoft.com/kb/193625
  if (req.status == 12029)
  { return new Error("The attempt to connect to the server failed."); }
  if (req.status == 12007)
  { return new Error("The server name could not be resolved."); }
  var errorText;
  try
        { errorText = JSON.parse(req.responseText).error.message.value; }
  catch (e)
        { errorText = req.responseText }

  return new Error("Error : " +
        req.status + ": " +
        req.statusText + ": " + errorText);
 },
 _dateReviver: function (key, value) {
  ///<summary>
  /// Private function to convert matching string values to Date objects.
  ///</summary>
  ///<param name="key" type="String">
  /// The key used to identify the object property
  ///</param>
  ///<param name="value" type="String">
  /// The string value representing a date
  ///</param>
  var a;
  if (typeof value === 'string') {
   a = /Date\(([-+]?\d+)\)/.exec(value);
   if (a) {
    return new Date(parseInt(value.replace("/Date(", "").replace(")/", ""), 10));
   }
  }
  return value;
 },
 _parameterCheck: function (parameter, message) {
  ///<summary>
  /// Private function used to check whether required parameters are null or undefined
  ///</summary>
  ///<param name="parameter" type="Object">
  /// The parameter to check;
  ///</param>
  ///<param name="message" type="String">
  /// The error message text to include when the error is thrown.
  ///</param>
  if ((typeof parameter === "undefined") || parameter === null) {
   throw new Error(message);
  }
 },
 _stringParameterCheck: function (parameter, message) {
  ///<summary>
  /// Private function used to check whether required parameters are null or undefined
  ///</summary>
  ///<param name="parameter" type="String">
  /// The string parameter to check;
  ///</param>
  ///<param name="message" type="String">
  /// The error message text to include when the error is thrown.
  ///</param>
  if (typeof parameter != "string") {
   throw new Error(message);
  }
 },
 _callbackParameterCheck: function (callbackParameter, message) {
  ///<summary>
  /// Private function used to check whether required callback parameters are functions
  ///</summary>
  ///<param name="callbackParameter" type="Function">
  /// The callback parameter to check;
  ///</param>
  ///<param name="message" type="String">
  /// The error message text to include when the error is thrown.
  ///</param>
  if (typeof callbackParameter != "function") {
   throw new Error(message);
  }
 },
 createRecord: function (object, type, successCallback, errorCallback) {
  ///<summary>
  /// Sends an asynchronous request to create a new record.
  ///</summary>
  ///<param name="object" type="Object">
  /// A JavaScript object with properties corresponding to the Schema name of
  /// entity attributes that are valid for create operations.
  ///</param>
  ///<param name="type" type="String">
  /// The Schema Name of the Entity type record to create.
  /// For an Account record, use "Account"
  ///</param>
  ///<param name="successCallback" type="Function">
  /// The function that will be passed through and be called by a successful response. 
  /// This function can accept the returned record as a parameter.
  /// </param>
  ///<param name="errorCallback" type="Function">
  /// The function that will be passed through and be called by a failed response. 
  /// This function must accept an Error object as a parameter.
  /// </param>
  this._parameterCheck(object, "SDK.REST.createRecord requires the object parameter.");
  this._stringParameterCheck(type, "SDK.REST.createRecord requires the type parameter is a string.");
  this._callbackParameterCheck(successCallback, "SDK.REST.createRecord requires the successCallback is a function.");
  this._callbackParameterCheck(errorCallback, "SDK.REST.createRecord requires the errorCallback is a function.");
  var req = new XMLHttpRequest();
  req.open("POST", encodeURI(this._ODataPath() + type + "Set"), true);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.onreadystatechange = function () {
   if (this.readyState == 4 /* complete */) {
    req.onreadystatechange = null;
    if (this.status == 201) {
     successCallback(JSON.parse(this.responseText, SDK.REST._dateReviver).d);
    }
    else {
     errorCallback(SDK.REST._errorHandler(this));
    }
   }
  };
  req.send(JSON.stringify(object));
 },
 retrieveRecord: function (id, type, select, expand, successCallback, errorCallback) {
  ///<summary>
  /// Sends an asynchronous request to retrieve a record.
  ///</summary>
  ///<param name="id" type="String">
  /// A String representing the GUID value for the record to retrieve.
  ///</param>
  ///<param name="type" type="String">
  /// The Schema Name of the Entity type record to retrieve.
  /// For an Account record, use "Account"
  ///</param>
  ///<param name="select" type="String">
  /// A String representing the $select OData System Query Option to control which
  /// attributes will be returned. This is a comma separated list of Attribute names that are valid for retrieve.
  /// If null all properties for the record will be returned
  ///</param>
  ///<param name="expand" type="String">
  /// A String representing the $expand OData System Query Option value to control which
  /// related records are also returned. This is a comma separated list of of up to 6 entity relationship names
  /// If null no expanded related records will be returned.
  ///</param>
  ///<param name="successCallback" type="Function">
  /// The function that will be passed through and be called by a successful response. 
  /// This function must accept the returned record as a parameter.
  /// </param>
  ///<param name="errorCallback" type="Function">
  /// The function that will be passed through and be called by a failed response. 
  /// This function must accept an Error object as a parameter.
  /// </param>
  this._stringParameterCheck(id, "SDK.REST.retrieveRecord requires the id parameter is a string.");
  this._stringParameterCheck(type, "SDK.REST.retrieveRecord requires the type parameter is a string.");
  if (select != null)
   this._stringParameterCheck(select, "SDK.REST.retrieveRecord requires the select parameter is a string.");
  if (expand != null)
   this._stringParameterCheck(expand, "SDK.REST.retrieveRecord requires the expand parameter is a string.");
  this._callbackParameterCheck(successCallback, "SDK.REST.retrieveRecord requires the successCallback parameter is a function.");
  this._callbackParameterCheck(errorCallback, "SDK.REST.retrieveRecord requires the errorCallback parameter is a function.");

  var systemQueryOptions = "";

  if (select != null || expand != null) {
   systemQueryOptions = "?";
   if (select != null) {
    var selectString = "$select=" + select;
    if (expand != null) {
     selectString = selectString + "," + expand;
    }
    systemQueryOptions = systemQueryOptions + selectString;
   }
   if (expand != null) {
    systemQueryOptions = systemQueryOptions + "&$expand=" + expand;
   }
  }


  var req = new XMLHttpRequest();
  req.open("GET", encodeURI(this._ODataPath() + type + "Set(guid'" + id + "')" + systemQueryOptions), true);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.onreadystatechange = function () {
   if (this.readyState == 4 /* complete */) {
    req.onreadystatechange = null;
    if (this.status == 200) {
     successCallback(JSON.parse(this.responseText, SDK.REST._dateReviver).d);
    }
    else {
     errorCallback(SDK.REST._errorHandler(this));
    }
   }
  };
  req.send();
 },
 updateRecord: function (id, object, type, successCallback, errorCallback) {
  ///<summary>
  /// Sends an asynchronous request to update a record.
  ///</summary>
  ///<param name="id" type="String">
  /// A String representing the GUID value for the record to retrieve.
  ///</param>
  ///<param name="object" type="Object">
  /// A JavaScript object with properties corresponding to the Schema Names for
  /// entity attributes that are valid for update operations.
  ///</param>
  ///<param name="type" type="String">
  /// The Schema Name of the Entity type record to retrieve.
  /// For an Account record, use "Account"
  ///</param>
  ///<param name="successCallback" type="Function">
  /// The function that will be passed through and be called by a successful response. 
  /// Nothing will be returned to this function.
  /// </param>
  ///<param name="errorCallback" type="Function">
  /// The function that will be passed through and be called by a failed response. 
  /// This function must accept an Error object as a parameter.
  /// </param>
  this._stringParameterCheck(id, "SDK.REST.updateRecord requires the id parameter.");
  this._parameterCheck(object, "SDK.REST.updateRecord requires the object parameter.");
  this._stringParameterCheck(type, "SDK.REST.updateRecord requires the type parameter.");
  this._callbackParameterCheck(successCallback, "SDK.REST.updateRecord requires the successCallback is a function.");
  this._callbackParameterCheck(errorCallback, "SDK.REST.updateRecord requires the errorCallback is a function.");
  var req = new XMLHttpRequest();

  req.open("POST", encodeURI(this._ODataPath() + type + "Set(guid'" + id + "')"), true);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("X-HTTP-Method", "MERGE");
  req.onreadystatechange = function () {
   if (this.readyState == 4 /* complete */) {
    req.onreadystatechange = null;
    if (this.status == 204 || this.status == 1223) {
     successCallback();
    }
    else {
     errorCallback(SDK.REST._errorHandler(this));
    }
   }
  };
  req.send(JSON.stringify(object));
 },
 deleteRecord: function (id, type, successCallback, errorCallback) {
  ///<summary>
  /// Sends an asynchronous request to delete a record.
  ///</summary>
  ///<param name="id" type="String">
  /// A String representing the GUID value for the record to delete.
  ///</param>
  ///<param name="type" type="String">
  /// The Schema Name of the Entity type record to delete.
  /// For an Account record, use "Account"
  ///</param>
  ///<param name="successCallback" type="Function">
  /// The function that will be passed through and be called by a successful response. 
  /// Nothing will be returned to this function.
  /// </param>
  ///<param name="errorCallback" type="Function">
  /// The function that will be passed through and be called by a failed response. 
  /// This function must accept an Error object as a parameter.
  /// </param>
  this._stringParameterCheck(id, "SDK.REST.deleteRecord requires the id parameter.");
  this._stringParameterCheck(type, "SDK.REST.deleteRecord requires the type parameter.");
  this._callbackParameterCheck(successCallback, "SDK.REST.deleteRecord requires the successCallback is a function.");
  this._callbackParameterCheck(errorCallback, "SDK.REST.deleteRecord requires the errorCallback is a function.");
  var req = new XMLHttpRequest();
  req.open("POST", encodeURI(this._ODataPath() + type + "Set(guid'" + id + "')"), true);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("X-HTTP-Method", "DELETE");
  req.onreadystatechange = function () {

   if (this.readyState == 4 /* complete */) {
    req.onreadystatechange = null;
    if (this.status == 204 || this.status == 1223) {
     successCallback();
    }
    else {
     errorCallback(SDK.REST._errorHandler(this));
    }
   }
  };
  req.send();

 },
 retrieveMultipleRecords: function (type, options, successCallback, errorCallback, OnComplete) {
  ///<summary>
  /// Sends an asynchronous request to retrieve records.
  ///</summary>
  ///<param name="type" type="String">
  /// The Schema Name of the Entity type record to retrieve.
  /// For an Account record, use "Account"
  ///</param>
  ///<param name="options" type="String">
  /// A String representing the OData System Query Options to control the data returned
  ///</param>
  ///<param name="successCallback" type="Function">
  /// The function that will be passed through and be called for each page of records returned.
  /// Each page is 50 records. If you expect that more than one page of records will be returned,
  /// this function should loop through the results and push the records into an array outside of the function.
  /// Use the OnComplete event handler to know when all the records have been processed.
  /// </param>
  ///<param name="errorCallback" type="Function">
  /// The function that will be passed through and be called by a failed response. 
  /// This function must accept an Error object as a parameter.
  /// </param>
  ///<param name="OnComplete" type="Function">
  /// The function that will be called when all the requested records have been returned.
  /// No parameters are passed to this function.
  /// </param>
  this._stringParameterCheck(type, "SDK.REST.retrieveMultipleRecords requires the type parameter is a string.");
  if (options != null)
   this._stringParameterCheck(options, "SDK.REST.retrieveMultipleRecords requires the options parameter is a string.");
  this._callbackParameterCheck(successCallback, "SDK.REST.retrieveMultipleRecords requires the successCallback parameter is a function.");
  this._callbackParameterCheck(errorCallback, "SDK.REST.retrieveMultipleRecords requires the errorCallback parameter is a function.");
  this._callbackParameterCheck(OnComplete, "SDK.REST.retrieveMultipleRecords requires the OnComplete parameter is a function.");

  var optionsString;
  if (options != null) {
   if (options.charAt(0) != "?") {
    optionsString = "?" + options;
   }
   else
   { optionsString = options; }
  }
  var req = new XMLHttpRequest();
  req.open("GET", this._ODataPath() + type + "Set" + optionsString, true);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.onreadystatechange = function () {
   if (this.readyState == 4 /* complete */) {
    req.onreadystatechange = null;
    if (this.status == 200) {
     var returned = JSON.parse(this.responseText, SDK.REST._dateReviver).d;
     successCallback(returned.results);
     if (returned.__next != null) {
      var queryOptions = returned.__next.substring((SDK.REST._ODataPath() + type + "Set").length);
      SDK.REST.retrieveMultipleRecords(type, queryOptions, successCallback, errorCallback, OnComplete);
     }
     else
     { OnComplete(); }
    }
    else {
     errorCallback(SDK.REST._errorHandler(this));
    }
   }
  };
  req.send();
 },
 associateRecords: function (parentId, parentType, relationshipName, childId, childType, successCallback, errorCallback) {
  this._stringParameterCheck(parentId, "SDK.REST.associateRecords requires the parentId parameter is a string.");
  ///<param name="parentId" type="String">
  /// The Id of the record to be the parent record in the relationship
  /// </param>
  ///<param name="parentType" type="String">
  /// The Schema Name of the Entity type for the parent record.
  /// For an Account record, use "Account"
  /// </param>
  ///<param name="relationshipName" type="String">
  /// The Schema Name of the Entity Relationship to use to associate the records.
  /// To associate account records as a Parent account, use "Referencedaccount_parent_account"
  /// </param>
  ///<param name="childId" type="String">
  /// The Id of the record to be the child record in the relationship
  /// </param>
  ///<param name="childType" type="String">
  /// The Schema Name of the Entity type for the child record.
  /// For an Account record, use "Account"
  /// </param>
  ///<param name="successCallback" type="Function">
  /// The function that will be passed through and be called by a successful response. 
  /// Nothing will be returned to this function.
  /// </param>
  ///<param name="errorCallback" type="Function">
  /// The function that will be passed through and be called by a failed response. 
  /// This function must accept an Error object as a parameter.
  /// </param>
  this._stringParameterCheck(parentType, "SDK.REST.associateRecords requires the parentType parameter is a string.");
  this._stringParameterCheck(relationshipName, "SDK.REST.associateRecords requires the relationshipName parameter is a string.");
  this._stringParameterCheck(childId, "SDK.REST.associateRecords requires the childId parameter is a string.");
  this._stringParameterCheck(childType, "SDK.REST.associateRecords requires the childType parameter is a string.");
  this._callbackParameterCheck(successCallback, "SDK.REST.associateRecords requires the successCallback parameter is a function.");
  this._callbackParameterCheck(errorCallback, "SDK.REST.associateRecords requires the errorCallback parameter is a function.");

  var req = new XMLHttpRequest();
  req.open("POST", encodeURI(this._ODataPath() + parentType + "Set(guid'" + parentId + "')/$links/" + relationshipName), true);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.onreadystatechange = function () {
   if (this.readyState == 4 /* complete */) {
    req.onreadystatechange = null;
    if (this.status == 204 || this.status == 1223) {
     successCallback();
    }
    else {
     errorCallback(SDK.REST._errorHandler(this));
    }
   }
  };
  var childEntityReference = {}
  childEntityReference.uri = this._ODataPath() + "/" + childType + "Set(guid'" + childId + "')";
  req.send(JSON.stringify(childEntityReference));
 },
 disassociateRecords: function (parentId, parentType, relationshipName, childId, successCallback, errorCallback) {
  this._stringParameterCheck(parentId, "SDK.REST.disassociateRecords requires the parentId parameter is a string.");
  ///<param name="parentId" type="String">
  /// The Id of the record to be the parent record in the relationship
  /// </param>
  ///<param name="parentType" type="String">
  /// The Schema Name of the Entity type for the parent record.
  /// For an Account record, use "Account"
  /// </param>
  ///<param name="relationshipName" type="String">
  /// The Schema Name of the Entity Relationship to use to disassociate the records.
  /// To disassociate account records as a Parent account, use "Referencedaccount_parent_account"
  /// </param>
  ///<param name="childId" type="String">
  /// The Id of the record to be disassociated as the child record in the relationship
  /// </param>
  ///<param name="successCallback" type="Function">
  /// The function that will be passed through and be called by a successful response. 
  /// Nothing will be returned to this function.
  /// </param>
  ///<param name="errorCallback" type="Function">
  /// The function that will be passed through and be called by a failed response. 
  /// This function must accept an Error object as a parameter.
  /// </param>
  this._stringParameterCheck(parentType, "SDK.REST.disassociateRecords requires the parentType parameter is a string.");
  this._stringParameterCheck(relationshipName, "SDK.REST.disassociateRecords requires the relationshipName parameter is a string.");
  this._stringParameterCheck(childId, "SDK.REST.disassociateRecords requires the childId parameter is a string.");
  this._callbackParameterCheck(successCallback, "SDK.REST.disassociateRecords requires the successCallback parameter is a function.");
  this._callbackParameterCheck(errorCallback, "SDK.REST.disassociateRecords requires the errorCallback parameter is a function.");

  var req = new XMLHttpRequest();
  req.open("POST", encodeURI(this._ODataPath() + parentType + "Set(guid'" + parentId + "')/$links/" + relationshipName + "(guid'" + childId + "')"), true);
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("X-HTTP-Method", "DELETE");
  req.onreadystatechange = function () {
   if (this.readyState == 4 /* complete */) {
    req.onreadystatechange = null;
    if (this.status == 204 || this.status == 1223) {
     successCallback();
    }
    else {
     errorCallback(SDK.REST._errorHandler(this));
    }
   }
  };
  req.send();
 },
 __namespace: true
};
// </snippetSDKRESTJS>
