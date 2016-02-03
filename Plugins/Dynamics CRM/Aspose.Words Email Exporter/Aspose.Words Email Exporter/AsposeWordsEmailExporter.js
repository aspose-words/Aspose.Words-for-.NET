function OpenEmailExporterPopup() {
    if (Xrm.Page.ui.getFormType() == 1)
    {
        alert("Save the record first");
        return;
    }
    var query = "$select=aspose_Parameter,aspose_Value&$filter=aspose_name eq 'EmailExporter'";
    SDK.REST.retrieveMultipleRecords("aspose_configuration", query, retrieveConfigSuccessfullPopup, function (error) { ShowError("Error retrieving configuration: " + error.message); }, function () { });
}
function retrieveConfigSuccessfullPopup(results) {
    if (results && results.length > 0) {
        for (var i in results) {
            var AsposeConfig = results[i];
            if (AsposeConfig.aspose_Parameter && AsposeConfig.aspose_Value) {
                if (AsposeConfig.aspose_Parameter == "URL") {
                    var URL = AsposeConfig.aspose_Value;
                    var id = Xrm.Page.data.entity.getId();
                    var orgname = Xrm.Page.context.getOrgUniqueName();
                    var typename = Xrm.Page.data.entity.getEntityName();
                    var FullURL = URL + "AsposeWordsEmailExporter.aspx?id=" + id + "&orgname=" + orgname + "&typename=" + typename;
                    window.open(FullURL, null, "width=400, height=400", true);
                }
            }
        }
    }
}