function GetTokenForEmail() {
    var webResourceName = "/webresources/aspose_BarCodeGenerator/GetToken.html";
    var DialogOption = new parent.Xrm.DialogOptions;
    DialogOption.width = 400;
    DialogOption.height = 400;
    Xrm.Internal.openDialog(webResourceName, DialogOption, null, null, function (returnValue) {
        var EmailBody = Xrm.Page.getAttribute("description").getValue();
        EmailBody += "<p>" + returnValue + "</p>";
        Xrm.Page.getAttribute("description").setValue(EmailBody);
    });
    
}
function LoadConfigurations() {
    var DD_BarCodeConfigurations = document.getElementById("DD_BarCodeConfigurations");
    removeOptions(DD_BarCodeConfigurations);
    var option = document.createElement("option");
    option.text = "Loading...";
    DD_BarCodeConfigurations.add(option);
    var query = "$select=*";
    SDK.REST.retrieveMultipleRecords("aspose_barcodeconfiguration", query, retrieveConfigSuccessfull, function (error) { ShowError("Error retrieving templates list: " + error.message); }, function () { });
}
function retrieveConfigSuccessfull(BarCodeConfigs) {
    if (BarCodeConfigs.length > 0) {
        var DD_BarCodeConfigurations = document.getElementById("DD_BarCodeConfigurations");
        removeOptions(DD_BarCodeConfigurations);
        for (var i in BarCodeConfigs) {
            var option = document.createElement("option");
            if (BarCodeConfigs[i].aspose_name)
                option.text = BarCodeConfigs[i].aspose_name;
            else
                option.text = "--";
            option.value = BarCodeConfigs[i].aspose_barcodeconfigurationId;
            DD_BarCodeConfigurations.add(option);
        }
    }
    else {
        var DD_Templates = document.getElementById("DD_Templates");
        removeOptions(DD_Templates);
    }
    BarCodeSelected();
}
function removeOptions(selectbox) {
    for (var i = selectbox.options.length - 1; i >= 0; i--) {
        selectbox.remove(i);
    }
}
function BarCodeSelected() {
    var DD_BarCodeConfigurations = document.getElementById("DD_BarCodeConfigurations");
    if (DD_BarCodeConfigurations.selectedIndex == -1 || DD_BarCodeConfigurations.options[DD_BarCodeConfigurations.selectedIndex].value == "") {
        return;
    }
    var SelectedBarCodeId = DD_BarCodeConfigurations.options[DD_BarCodeConfigurations.selectedIndex].value;
    SelectedBarCodeId = SelectedBarCodeId.replace(/[{}]/g, "");
    var TXT_Token = document.getElementById("TXT_Token");
    TXT_Token.value = "[AsposeBarCode{" + SelectedBarCodeId + "}]";
}
function Insert()
{
    var TXT_Token = document.getElementById("TXT_Token");
    Mscrm.Utilities.setReturnValue(TXT_Token.value);
    closeWindow(true);
}