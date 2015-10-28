function CKEditorDocumentReady() {
    var Xrm = parent.Xrm;
    CKEDITOR.replace("editor1");

    var Contents = Xrm.Page.getAttribute("aspose_body").getValue(); //Load data from CRM to CK Editor
    if (Contents) //if data already exist
        CKEDITOR.instances.editor1.setData(Contents);

    CKEDITOR.instances.editor1.on('blur', function () {
        // Call this function when CK Editor loose focus, to set the update value in CRM
        var value = CKEDITOR.instances.editor1.getData();
        //value = UpdateHTML(value);
        if (value)
            Xrm.Page.getAttribute("aspose_body").setValue(value);
        else
            Xrm.Page.getAttribute("aspose_body").setValue("");
    });
}


function AddButton() { // Add Button on CK Editor Ribbon
    CKEDITOR.on('instanceCreated', function (ev) {
        var editor = ev.editor;
        editor.on('pluginsLoaded', function () {
            editor.addCommand('AddFields', {
                exec: function (editor) {
                    var url = "/webresources/Aspose_AsposeQuoteGenerator/InsertMergeField.html";
                    var DialogOption = new parent.Xrm.DialogOptions;
                    DialogOption.width = 400;
                    DialogOption.height = 400;

                    parent.Xrm.Internal.openDialog(url,
                                            DialogOption,
                                            null, null,
                                            function (returnValue) {
                                                editor.insertText(returnValue);
                                            });
                },
                canUndo: false
            });
            editor.ui.add('AddFields', CKEDITOR.UI_BUTTON, {
                label: 'Insert CRM Field',
                command: 'AddFields'
            });
        });
    });
}
function LoadFields() {
    //debugger;
    // if the radio button is changed, fill dropdown
    var DD_Fields = document.getElementById("DD_Fields");
    removeOptions(DD_Fields);
    var option = document.createElement("option");
    option.text = "Loading...";
    DD_Fields.add(option); //Load Quote Fields if not already retrieved and load in dropdown
    SDK.Metadata.RetrieveEntity(SDK.Metadata.EntityFilters.Attributes,
       "Quote",
       null,
       false,
       function (entityMetadata) { successRetrieveEntityAttributes("Quote", entityMetadata); },
       errorRetrieveEntity);
}
function successRetrieveEntityAttributes(logicalName, entityMetadata) {
    entityMetadata.Attributes.sort(function (a, b) {
        if (a.LogicalName < b.LogicalName)
        { return -1 }
        if (a.LogicalName > b.LogicalName)
        { return 1 }
        return 0;
    });
   var fieldsList = entityMetadata.Attributes; // Successfully retrieved Quote Fields from MetaData

    var DD_Fields = document.getElementById("DD_Fields");
    removeOptions(DD_Fields);
    for (var i in fieldsList) {
        if (AddThisField(fieldsList[i])) {
            var option = document.createElement("option");
            option.text = fieldsList[i].DisplayName.UserLocalizedLabel.Label + " {" + fieldsList[i].SchemaName + "}";
            option.value = fieldsList[i].SchemaName;
            DD_Fields.add(option);
        }
    }
}
function errorRetrieveEntity(error) {
    alert(error.message);
}
function removeOptions(selectbox) { // Empty dropdown before adding new values
    for (var i = selectbox.options.length - 1; i >= 0; i--) {
        selectbox.remove(i);
    }
}
function AddThisField(Field) {
    var skippedFields = ["_Base", "CreatedOnBehalfBy", "CustomerIdType",
        "OwningBusinessUnit", "OwningTeam", "OwningUser", "UniqueDscId", "UTCConversionTimeZoneCode",
        "QuoteId", "TimeZoneRuleVersionNumber", "ModifiedOnBehalfBy", "OverriddenCreatedOn", "ProcessId",
        "StageId", "CampaignId"];
    if (Field.DisplayName.UserLocalizedLabel && Field.DisplayName.UserLocalizedLabel.Label && Field.SchemaName) {
        for (var i in skippedFields) {
            if (Field.SchemaName.toLowerCase().indexOf(skippedFields[i].toLowerCase()) >= 0)
                return false;
        }
        return true;
    }
    else
        return false;
}
function Insert() {
    // Insert button is pressed (It will insert the selected field from the dropdown to the CK Editor)
    var DD_Fields = document.getElementById("DD_Fields");
    var SelectedAttribute = DD_Fields.options[DD_Fields.selectedIndex].value;
    Mscrm.Utilities.setReturnValue("<MERGEFIELD>" + SelectedAttribute + "</MERGEFIELD>");
    closeWindow(true);
}