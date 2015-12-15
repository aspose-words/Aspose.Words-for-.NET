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

//<snippetSDK.Metadata.js>

SDK = window.SDK || { __namespace: true };
SDK.Metadata = SDK.Metadata || { __namespace: true };

(function () {


 this.RetrieveAllEntities = function (EntityFilters, RetrieveAsIfPublished, successCallBack, errorCallBack, passThroughObject) {
  ///<summary>
  /// Sends an asynchronous RetrieveAllEntities Request to retrieve all entities in the system
  ///</summary>
  ///<returns>entityMetadataCollection</returns>
  ///<param name="EntityFilters" type="Number">
  /// SDK.Metadata.EntityFilters provides an enumeration for the filters available to filter which data is retrieved.
  /// Include only those elements of the entity you want to retrieve. Retrieving all parts of all entitities may take significant time.
  ///</param>
  ///<param name="RetrieveAsIfPublished" type="Boolean">
  /// Sets whether to retrieve the metadata that has not been published.
  ///</param>
  ///<param name="successCallBack" type="Function">
  /// The function that will be passed through and be called by a successful response.
  /// This function must accept the entityMetadataCollection as a parameter.
  ///</param>
  ///<param name="errorCallBack" type="Function">
  /// The function that will be passed through and be called by a failed response.
  /// This function must accept an Error object as a parameter.
  ///</param>
  ///<param name="passThroughObject" optional="true"  type="Object">
  /// An Object that will be passed through to as the second parameter to the successCallBack.
  ///</param>
  if ((typeof EntityFilters != "number") || (EntityFilters < 1 || EntityFilters > 15))
  { throw new Error("SDK.Metadata.RetrieveAllEntities EntityFilters must be a SDK.Metadata.EntityFilters value."); }
  if (typeof RetrieveAsIfPublished != "boolean")
  { throw new Error("SDK.Metadata.RetrieveAllEntities RetrieveAsIfPublished must be a boolean value."); }
  if (typeof successCallBack != "function")
  { throw new Error("SDK.Metadata.RetrieveAllEntities successCallBack must be a function."); }
  if (typeof errorCallBack != "function")
  { throw new Error("SDK.Metadata.RetrieveAllEntities errorCallBack must be a function."); }

  var entityFiltersValue = _evaluateEntityFilters(EntityFilters);

  var request = [
  "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\">",
    //Allows retrieval if ImageAttributeMetadata objects
  "<soapenv:Header><a:SdkClientVersion xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">7.0</a:SdkClientVersion></soapenv:Header>",
   "<soapenv:Body>",
    "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">",
     "<request i:type=\"a:RetrieveAllEntitiesRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">",
      "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>EntityFilters</b:key>",
        "<b:value i:type=\"c:EntityFilters\" xmlns:c=\"http://schemas.microsoft.com/xrm/2011/Metadata\">" + _xmlEncode(entityFiltersValue) + "</b:value>",
       "</a:KeyValuePairOfstringanyType>",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>RetrieveAsIfPublished</b:key>",
        "<b:value i:type=\"c:boolean\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + _xmlEncode(RetrieveAsIfPublished.toString()) + "</b:value>",
       "</a:KeyValuePairOfstringanyType>",
      "</a:Parameters>",
      "<a:RequestId i:nil=\"true\" />",
      "<a:RequestName>RetrieveAllEntities</a:RequestName>",
    "</request>",
   "</Execute>",
  "</soapenv:Body>",
  "</soapenv:Envelope>"].join("");
  var req = new XMLHttpRequest();
  req.open("POST", _getUrl() + "/XRMServices/2011/Organization.svc/web", true);
  try { req.responseType = 'msxml-document'} catch(e){}
  req.setRequestHeader("Accept", "application/xml, text/xml, */*");
  req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
  req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");
  req.onreadystatechange = function () {
   if (req.readyState == 4 /* complete */) {
   req.onreadystatechange = null; //Addresses potential memory leak issue with IE
    if (req.status == 200) {
     //Success				
     var doc = req.responseXML;
     try{_setSelectionNamespaces(doc);}catch(e){}
     var entityMetadataNodes = _selectNodes(doc, "//c:EntityMetadata");
     var entityMetadataCollection = [];
     for (var i = 0; i < entityMetadataNodes.length; i++) {
      var a = _objectifyNode(entityMetadataNodes[i]);
      a._type = "EntityMetadata";
      entityMetadataCollection.push(a);
     }

     successCallBack(entityMetadataCollection, passThroughObject);
    }
    else {
     errorCallBack(_getError(req));
    }
   }
  };
  req.send(request);

 };

 this.RetrieveEntity = function (EntityFilters, LogicalName, MetadataId, RetrieveAsIfPublished, successCallBack, errorCallBack, passThroughObject) {
  ///<summary>
  /// Sends an asynchronous RetrieveEntity Request to retrieve a specific entity
  ///</summary>
  ///<returns>entityMetadata</returns>
  ///<param name="EntityFilters" type="Number">
  /// SDK.Metadata.EntityFilters provides an enumeration for the filters available to filter which data is retrieved.
  /// Include only those elements of the entity you want to retrieve. Retrieving all parts of all entitities may take significant time.
  ///</param>
  ///<param name="LogicalName" optional="true" type="String">
  /// The logical name of the entity requested. A null value may be used if a MetadataId is provided.
  ///</param>
  ///<param name="MetadataId" optional="true" type="String">
  /// A null value or an empty guid may be passed if a LogicalName is provided.
  ///</param>
  ///<param name="RetrieveAsIfPublished" type="Boolean">
  /// Sets whether to retrieve the metadata that has not been published.
  ///</param>
  ///<param name="successCallBack" type="Function">
  /// The function that will be passed through and be called by a successful response.
  /// This function must accept the entityMetadata as a parameter.
  ///</param>
  ///<param name="errorCallBack" type="Function">
  /// The function that will be passed through and be called by a failed response.
  /// This function must accept an Error object as a parameter.
  ///</param>
  ///<param name="passThroughObject" optional="true"  type="Object">
  /// An Object that will be passed through to as the second parameter to the successCallBack.
  ///</param>
  if ((typeof EntityFilters != "number") || (EntityFilters < 1 || EntityFilters > 15))
  { throw new Error("SDK.Metadata.RetrieveEntity EntityFilters must be a SDK.Metadata.EntityFilters value."); }
  if (LogicalName == null && MetadataId == null) {
   throw new Error("SDK.Metadata.RetrieveEntity requires either the LogicalName or MetadataId parameter not be null.");
  }
  if (LogicalName != null) {
   if (typeof LogicalName != "string")
   { throw new Error("SDK.Metadata.RetrieveEntity LogicalName must be a string value."); }
   MetadataId = "00000000-0000-0000-0000-000000000000";
  }
  if (MetadataId != null && LogicalName == null) {
   if (typeof MetadataId != "string")
   { throw new Error("SDK.Metadata.RetrieveEntity MetadataId must be a string value."); }
  }
  if (typeof RetrieveAsIfPublished != "boolean")
  { throw new Error("SDK.Metadata.RetrieveEntity RetrieveAsIfPublished must be a boolean value."); }
  if (typeof successCallBack != "function")
  { throw new Error("SDK.Metadata.RetrieveEntity successCallBack must be a function."); }
  if (typeof errorCallBack != "function")
  { throw new Error("SDK.Metadata.RetrieveEntity errorCallBack must be a function."); }
  var entityFiltersValue = _evaluateEntityFilters(EntityFilters);

  var entityLogicalNameValueNode = "";
  if (LogicalName == null)
  { entityLogicalNameValueNode = "<b:value i:nil=\"true\" />"; }
  else
  { entityLogicalNameValueNode = "<b:value i:type=\"c:string\"   xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + _xmlEncode(LogicalName.toLowerCase()) + "</b:value>"; }
  var request = [
  "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\">",
    //Allows retrieval if ImageAttributeMetadata objects
  "<soapenv:Header><a:SdkClientVersion xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">6.0</a:SdkClientVersion></soapenv:Header>",
   "<soapenv:Body>",
    "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">",
     "<request i:type=\"a:RetrieveEntityRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">",
      "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>EntityFilters</b:key>",
        "<b:value i:type=\"c:EntityFilters\" xmlns:c=\"http://schemas.microsoft.com/xrm/2011/Metadata\">" + _xmlEncode(entityFiltersValue) + "</b:value>",
       "</a:KeyValuePairOfstringanyType>",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>MetadataId</b:key>",
        "<b:value i:type=\"ser:guid\"  xmlns:ser=\"http://schemas.microsoft.com/2003/10/Serialization/\">" + _xmlEncode(MetadataId) + "</b:value>",
       "</a:KeyValuePairOfstringanyType>",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>RetrieveAsIfPublished</b:key>",
        "<b:value i:type=\"c:boolean\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + _xmlEncode(RetrieveAsIfPublished.toString()) + "</b:value>",
       "</a:KeyValuePairOfstringanyType>",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>LogicalName</b:key>",
         entityLogicalNameValueNode,
       "</a:KeyValuePairOfstringanyType>",
      "</a:Parameters>",
      "<a:RequestId i:nil=\"true\" />",
      "<a:RequestName>RetrieveEntity</a:RequestName>",
     "</request>",
    "</Execute>",
   "</soapenv:Body>",
  "</soapenv:Envelope>"].join("");
  var req = new XMLHttpRequest();
  req.open("POST", _getUrl() + "/XRMServices/2011/Organization.svc/web", true);
  try { req.responseType = 'msxml-document'} catch(e){}
  req.setRequestHeader("Accept", "application/xml, text/xml, */*");
  req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
  req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");
  req.onreadystatechange = function () {
   if (req.readyState == 4 /* complete */) {
   req.onreadystatechange = null; //Addresses potential memory leak issue with IE
    if (req.status == 200) {
     var doc = req.responseXML;
     try{_setSelectionNamespaces(doc);}catch(e){}
     var a = _objectifyNode(_selectSingleNode(doc, "//b:value"));
     a._type = "EntityMetadata";
     successCallBack(a, passThroughObject);
    }
    else {
     //Failure
     errorCallBack(_getError(req));
    }
   }

  };
  req.send(request);


 };

 this.RetrieveAttribute = function (EntityLogicalName, LogicalName, MetadataId, RetrieveAsIfPublished, successCallBack, errorCallBack, passThroughObject) {
  ///<summary>
  /// Sends an asynchronous RetrieveAttribute Request to retrieve a specific entity
  ///</summary>
  ///<returns>AttributeMetadata</returns>
  ///<param name="EntityLogicalName"  optional="true" type="String">
  /// The logical name of the entity for the attribute requested. A null value may be used if a MetadataId is provided.
  ///</param>
  ///<param name="LogicalName" optional="true" type="String">
  /// The logical name of the attribute requested. 
  ///</param>
  ///<param name="MetadataId" optional="true" type="String">
  /// A null value may be passed if an EntityLogicalName and LogicalName is provided.
  ///</param>
  ///<param name="RetrieveAsIfPublished" type="Boolean">
  /// Sets whether to retrieve the metadata that has not been published.
  ///</param>
  ///<param name="successCallBack" type="Function">
  /// The function that will be passed through and be called by a successful response.
  /// This function must accept the entityMetadata as a parameter.
  ///</param>
  ///<param name="errorCallBack" type="Function">
  /// The function that will be passed through and be called by a failed response.
  /// This function must accept an Error object as a parameter.
  ///</param>
  ///<param name="passThroughObject" optional="true"  type="Object">
  /// An Object that will be passed through to as the second parameter to the successCallBack.
  ///</param>
  if (EntityLogicalName == null && LogicalName == null && MetadataId == null) {
   throw new Error("SDK.Metadata.RetrieveAttribute requires either the EntityLogicalName and LogicalName  parameters or the MetadataId parameter not be null.");
  }
  if (MetadataId != null && EntityLogicalName == null && LogicalName == null) {
   if (typeof MetadataId != "string")
   { throw new Error("SDK.Metadata.RetrieveEntity MetadataId must be a string value."); }
  }
  else
  { MetadataId = "00000000-0000-0000-0000-000000000000"; }
  if (EntityLogicalName != null) {
   if (typeof EntityLogicalName != "string") {
    { throw new Error("SDK.Metadata.RetrieveAttribute EntityLogicalName must be a string value."); }
   }
  }
  if (LogicalName != null) {
   if (typeof LogicalName != "string") {
    { throw new Error("SDK.Metadata.RetrieveAttribute LogicalName must be a string value."); }
   }

  }
  if (typeof RetrieveAsIfPublished != "boolean")
  { throw new Error("SDK.Metadata.RetrieveAttribute RetrieveAsIfPublished must be a boolean value."); }
  if (typeof successCallBack != "function")
  { throw new Error("SDK.Metadata.RetrieveAttribute successCallBack must be a function."); }
  if (typeof errorCallBack != "function")
  { throw new Error("SDK.Metadata.RetrieveAttribute errorCallBack must be a function."); }

  var entityLogicalNameValueNode;
  if (EntityLogicalName == null) {
   entityLogicalNameValueNode = "<b:value i:nil=\"true\" />";
  }
  else {
   entityLogicalNameValueNode = "<b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + _xmlEncode(EntityLogicalName.toLowerCase()) + "</b:value>";
  }
  var logicalNameValueNode;
  if (LogicalName == null) {
   logicalNameValueNode = "<b:value i:nil=\"true\" />";
  }
  else {
   logicalNameValueNode = "<b:value i:type=\"c:string\"   xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + _xmlEncode(LogicalName.toLowerCase()) + "</b:value>";
  }
  var request = [
  "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\">",
  //Allows retrieval if ImageAttributeMetadata objects
  "<soapenv:Header><a:SdkClientVersion xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">6.0</a:SdkClientVersion></soapenv:Header>",
   "<soapenv:Body>",
    "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">",
     "<request i:type=\"a:RetrieveAttributeRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">",
      "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>EntityLogicalName</b:key>",
         entityLogicalNameValueNode,
       "</a:KeyValuePairOfstringanyType>",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>MetadataId</b:key>",
        "<b:value i:type=\"ser:guid\"  xmlns:ser=\"http://schemas.microsoft.com/2003/10/Serialization/\">" + _xmlEncode(MetadataId) + "</b:value>",
       "</a:KeyValuePairOfstringanyType>",
        "<a:KeyValuePairOfstringanyType>",
        "<b:key>RetrieveAsIfPublished</b:key>",
      "<b:value i:type=\"c:boolean\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + _xmlEncode(RetrieveAsIfPublished.toString()) + "</b:value>",
       "</a:KeyValuePairOfstringanyType>",
       "<a:KeyValuePairOfstringanyType>",
        "<b:key>LogicalName</b:key>",
         logicalNameValueNode,
       "</a:KeyValuePairOfstringanyType>",
      "</a:Parameters>",
      "<a:RequestId i:nil=\"true\" />",
      "<a:RequestName>RetrieveAttribute</a:RequestName>",
     "</request>",
    "</Execute>",
   "</soapenv:Body>",
  "</soapenv:Envelope>"].join("");
  var req = new XMLHttpRequest();
  req.open("POST", _getUrl() + "/XRMServices/2011/Organization.svc/web", true);
      try { req.responseType = 'msxml-document'} catch(e){}
  req.setRequestHeader("Accept", "application/xml, text/xml, */*");
  req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
  req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");
  req.onreadystatechange = function () {
   if (req.readyState == 4 /* complete */) {
   req.onreadystatechange = null; //Addresses potential memory leak issue with IE
    if (req.status == 200) {
     //Success
     var doc = req.responseXML;
     try{_setSelectionNamespaces(doc);}catch(e){}
     var a = _objectifyNode(_selectSingleNode(doc, "//b:value"));
    				
     successCallBack(a, passThroughObject);
    }
    else {
     //Failure
     errorCallBack(_getError(req));
    }
   }
  };
  req.send(request);


 };

 this.EntityFilters = function () {
  /// <summary>SDK.Metadata.EntityFilters enum summary</summary>
  /// <field name="Default" type="Number" static="true">enum field summary for Default</field>
  /// <field name="Entity" type="Number" static="true">enum field summary for Entity</field>
  /// <field name="Attributes" type="Number" static="true">enum field summary for Attributes</field>
  /// <field name="Privileges" type="Number" static="true">enum field summary for Privileges</field>
  /// <field name="Relationships" type="Number" static="true">enum field summary for Relationships</field>
  /// <field name="All" type="Number" static="true">enum field summary for All</field>
  throw new Error("Constructor not implemented this is a static enum.");
 };



 var _arrayElements = ["Attributes",
  "ManyToManyRelationships",
  "ManyToOneRelationships",
  "OneToManyRelationships",
  "Privileges",
  "LocalizedLabels",
  "Options",
  "Targets"];

  function _getError(resp) {
  ///<summary>
  /// Private function that attempts to parse errors related to connectivity or WCF faults.
  ///</summary>
  ///<param name="resp" type="XMLHttpRequest">
  /// The XMLHttpRequest representing failed response.
  ///</param>

  //Error descriptions come from http://support.microsoft.com/kb/193625
  if (resp.status == 12029)
  { return new Error("The attempt to connect to the server failed."); }
  if (resp.status == 12007)
  { return new Error("The server name could not be resolved."); }
  var faultXml = resp.responseXML;
  var errorMessage = "Unknown (unable to parse the fault)";
  if (typeof faultXml == "object") {

   var faultstring = null;
   var ErrorCode = null;

   var bodyNode = faultXml.firstChild.firstChild;

   //Retrieve the fault node
   for (var i = 0; i < bodyNode.childNodes.length; i++) {
    var node = bodyNode.childNodes[i];

    //NOTE: This comparison does not handle the case where the XML namespace changes
    if ("s:Fault" == node.nodeName) {
     for (var j = 0; j < node.childNodes.length; j++) {
      var testNode = node.childNodes[j];
      if ("faultstring" == testNode.nodeName) {
       faultstring = _getNodeText(testNode);
      }
      if ("detail" == testNode.nodeName) {
       for (var k = 0; k < testNode.childNodes.length; k++) {
        var orgServiceFault = testNode.childNodes[k];
        if ("OrganizationServiceFault" == orgServiceFault.nodeName) {
         for (var l = 0; l < orgServiceFault.childNodes.length; l++) {
          var ErrorCodeNode = orgServiceFault.childNodes[l];
          if ("ErrorCode" == ErrorCodeNode.nodeName) {
           ErrorCode = _getNodeText(ErrorCodeNode);
           break;
          }
         }
        }
       }

      }
     }
     break;
    }

   }
  }
  if (ErrorCode != null && faultstring != null) {
   errorMessage = "Error Code:" + ErrorCode + " Message: " + faultstring;
  }
  else {
   if (faultstring != null) {
    errorMessage = faultstring;
   }
  }
  return new Error(errorMessage);
 };

 function _Context() {
  var errorMessage = "Context is not available.";
  if (typeof GetGlobalContext != "undefined")
  { return GetGlobalContext(); }
  else {
   if (typeof Xrm != "undefined") {
    return Xrm.Page.context;
   }
   else
   { return new Error(errorMessage); }
  }
 };

 function _getUrl() {
     var url = _Context().getClientUrl();
  return url;
 };

 function _evaluateEntityFilters(EntityFilters) {
  var entityFilterArray = [];
  if ((1 & EntityFilters) == 1) {
   entityFilterArray.push("Entity");
  }
  if ((2 & EntityFilters) == 2) {
   entityFilterArray.push("Attributes");
  }
  if ((4 & EntityFilters) == 4) {
   entityFilterArray.push("Privileges");
  }
  if ((8 & EntityFilters) == 8) {
   entityFilterArray.push("Relationships");
  }
  return entityFilterArray.join(" ");
 };

 function _isMetadataArray(elementName) {
  for (var i = 0; i < _arrayElements.length; i++) {
   if (elementName == _arrayElements[i]) {
    return true;
   }
  }
  return false;
 };

 function _objectifyNode(node) {
  //Check for null
  if (node.attributes != null && node.attributes.length == 1) {
   if (node.attributes.getNamedItem("i:nil") != null && node.attributes.getNamedItem("i:nil").nodeValue == "true") {
    return null;
   }
  }

  //Check if it is a value
  if ((node.firstChild != null) && (node.firstChild.nodeType == 3)) {
   var nodeName = _getNodeName(node);

   switch (nodeName) {
    //Integer Values 
    case "ActivityTypeMask":
    case "ObjectTypeCode":
    case "ColumnNumber":
    case "DefaultFormValue":
    case "MaxValue":
    case "MinValue":
    case "MaxLength":
    case "Order":
    case "Precision":
    case "PrecisionSource":
    case "LanguageCode":
     return parseInt(node.firstChild.nodeValue, 10);
     // Boolean values
    case "AutoRouteToOwnerQueue":
    case "CanBeChanged":
    case "CanTriggerWorkflow":
    case "IsActivity":
    case "IsAIRUpdated":
    case "IsActivityParty":
    case "IsAvailableOffline":
    case "IsChildEntity":
    case "IsCustomEntity":
    case "IsCustomOptionSet":   
    case "IsDocumentManagementEnabled":
    case "IsEnabledForCharts":
    case "IsGlobal":
    case "IsImportable":
    case "IsIntersect":
    case "IsManaged":
    case "IsReadingPaneEnabled":
    case "IsValidForAdvancedFind":
    case "CanBeSecuredForCreate":
    case "CanBeSecuredForRead":
    case "CanBeSecuredForUpdate":
    case "IsCustomAttribute":
    case "IsManaged":
    case "IsPrimaryId":
    case "IsPrimaryName":
    case "IsSecured":
    case "IsValidForCreate":
    case "IsValidForRead":
    case "IsValidForUpdate":
    case "IsCustomRelationship":
    case "CanBeBasic":
    case "CanBeDeep":
    case "CanBeGlobal":
    case "CanBeLocal":
     return (node.firstChild.nodeValue == "true") ? true : false;
     //OptionMetadata.Value and BooleanManagedProperty.Value and AttributeRequiredLevelManagedProperty.Value
    case "Value":
     //BooleanManagedProperty.Value
     if ((node.firstChild.nodeValue == "true") || (node.firstChild.nodeValue == "false")) {
      return (node.firstChild.nodeValue == "true") ? true : false;
     }
     //AttributeRequiredLevelManagedProperty.Value
     if (
  (node.firstChild.nodeValue == "ApplicationRequired") ||
  (node.firstChild.nodeValue == "None") ||
  (node.firstChild.nodeValue == "Recommended") ||
  (node.firstChild.nodeValue == "SystemRequired")
  ) {
      return node.firstChild.nodeValue;
     }
     var numberValue = parseInt(node.firstChild.nodeValue, 10);
     if (isNaN(numberValue)) {
         //FormatName.Value
         return node.firstChild.nodeValue;
     }
     else {
         //OptionMetadata.Value
         return numberValue;
     }
     break;
    //String values 
    default:
     return node.firstChild.nodeValue;
   }

  }

  //Check if it is a known array
  if (_isMetadataArray(_getNodeName(node))) {
   var arrayValue = [];

   for (var i = 0; i < node.childNodes.length; i++) {
    var objectTypeName;
    if ((node.childNodes[i].attributes != null) && (node.childNodes[i].attributes.getNamedItem("i:type") != null)) {
     objectTypeName = node.childNodes[i].attributes.getNamedItem("i:type").nodeValue.split(":")[1];
    }
    else {

     objectTypeName = _getNodeName(node.childNodes[i]);
    }

    var b = _objectifyNode(node.childNodes[i]);
    b._type = objectTypeName;
    arrayValue.push(b);

   }

   return arrayValue;
  }

    //Null entity description labels are returned as <label/> - not using i:nil = true;
  if(node.childNodes.length == 0)
  {
  return null;
  }

  //Otherwise return an object
  var c = {};
  if (node.attributes.getNamedItem("i:type") != null) {
   c._type = node.attributes.getNamedItem("i:type").nodeValue.split(":")[1];
  }
  for (var i = 0; i < node.childNodes.length; i++) {
   if (node.childNodes[i].nodeType == 3) {
    c[_getNodeName(node.childNodes[i])] = node.childNodes[i].nodeValue;
   }
   else {
    c[_getNodeName(node.childNodes[i])] = _objectifyNode(node.childNodes[i]);
   }

  }
  return c;
 };

 function _selectNodes(node, XPathExpression) {
  if (typeof (node.selectNodes) != "undefined") {
   return node.selectNodes(XPathExpression);
  }
  else {
   var output = [];
   var XPathResults = node.evaluate(XPathExpression, node, _NSResolver, XPathResult.ANY_TYPE, null);
   var result = XPathResults.iterateNext();
   while (result) {
    output.push(result);
    result = XPathResults.iterateNext();
   }
   return output;
  }
 };

 function _selectSingleNode(node, xpathExpr) {
  if (typeof (node.selectSingleNode) != "undefined") {
   return node.selectSingleNode(xpathExpr);
  }
  else {
   var xpe = new XPathEvaluator();
   var xPathNode = xpe.evaluate(xpathExpr, node, _NSResolver, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
   return (xPathNode != null) ? xPathNode.singleNodeValue : null;
  }
 };

 function _selectSingleNodeText(node, xpathExpr) {
  var x = _selectSingleNode(node, xpathExpr);
  if (_isNodeNull(x))
  { return null; }
  if (typeof (x.text) != "undefined") {
   return x.text;
  }
  else {
   return x.textContent;
  }
 };

 function _getNodeText(node) {
  if (typeof (node.text) != "undefined") {
   return node.text;
  }
  else {
   return node.textContent;
  }
 };

 function _isNodeNull(node) {
  if (node == null)
  { return true; }
  if ((node.attributes.getNamedItem("i:nil") != null) && (node.attributes.getNamedItem("i:nil").value == "true"))
  { return true; }
  return false;
 };

 function _getNodeName(node) {
  if (typeof (node.baseName) != "undefined") {
   return node.baseName;
  }
  else {
   return node.localName;
  }
 };

 function _setSelectionNamespaces(doc)
 {
 var namespaces = [
 "xmlns:s='http://schemas.xmlsoap.org/soap/envelope/'",
 "xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts'",
 "xmlns:i='http://www.w3.org/2001/XMLSchema-instance'",
 "xmlns:b='http://schemas.datacontract.org/2004/07/System.Collections.Generic'",
 "xmlns:c='http://schemas.microsoft.com/xrm/2011/Metadata'" 
 ];
 doc.setProperty("SelectionNamespaces",namespaces.join(" "));
 
 }

 function _NSResolver(prefix) {
  var ns = {
   "s": "http://schemas.xmlsoap.org/soap/envelope/",
   "a": "http://schemas.microsoft.com/xrm/2011/Contracts",
   "i": "http://www.w3.org/2001/XMLSchema-instance",
   "b": "http://schemas.datacontract.org/2004/07/System.Collections.Generic",
   "c": "http://schemas.microsoft.com/xrm/2011/Metadata"
  };
  return ns[prefix] || null;
 };

 function _xmlEncode(strInput) {
  var c;
  var XmlEncode = '';
  if (strInput == null) {
   return null;
  }
  if (strInput == '') {
   return '';
  }
  for (var cnt = 0; cnt < strInput.length; cnt++) {
   c = strInput.charCodeAt(cnt);
   if (((c > 96) && (c < 123)) ||
			((c > 64) && (c < 91)) ||
			(c == 32) ||
			((c > 47) && (c < 58)) ||
			(c == 46) ||
			(c == 44) ||
			(c == 45) ||
			(c == 95)) {
    XmlEncode = XmlEncode + String.fromCharCode(c);
   }
   else {
    XmlEncode = XmlEncode + '&#' + c + ';';
   }
  }
  return XmlEncode;
 };

}).call(SDK.Metadata);

//SDK.Metadata.EntityFilters
// this enum is written this way to enable Visual Studio IntelliSense
SDK.Metadata.EntityFilters.prototype = {
 Default: 1,
 Entity: 1,
 Attributes: 2,
 Privileges: 4,
 Relationships: 8,
 All: 15
};
SDK.Metadata.EntityFilters.Default = 1;
SDK.Metadata.EntityFilters.Entity = 1;
SDK.Metadata.EntityFilters.Attributes = 2;
SDK.Metadata.EntityFilters.Privileges = 4;
SDK.Metadata.EntityFilters.Relationships = 8;
SDK.Metadata.EntityFilters.All = 15;
SDK.Metadata.EntityFilters.__enum = true;
SDK.Metadata.EntityFilters.__flags = true;


//</snippetSDK.Metadata.js>


