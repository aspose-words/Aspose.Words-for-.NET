<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="View.ascx.cs" Inherits="Aspose.Modules.AsposeDotNetNukeContentExport.View" %>

<script type="text/javascript">

    String.prototype.replaceAll = function (find, replace) {
        var str = this;
        return str.replace(new RegExp(find, 'g'), replace);
    };

    function ButtonClicked() {
        var PanesDropDownList = document.getElementById("<%=PanesDropDownList.ClientID%>");
        var selectedDropDownValue = PanesDropDownList.options[PanesDropDownList.selectedIndex].value;

        if (selectedDropDownValue == "dnn_full_page") {
            document.getElementById("PageSourceHiddenField").value = document.body.innerHTML.replaceAll("<", "#l#").replaceAll(">", "#g#");
        }
        else {
            document.getElementById("PageSourceHiddenField").value = document.getElementById(selectedDropDownValue).innerHTML.replaceAll("<", "#l#").replaceAll(">", "#g#");
        }
        return true;
    }
</script>


<div class="exportButton">
    <asp:HiddenField ID="PageSourceHiddenField" ClientIDMode="Static" runat="server" />
    <asp:DropDownList ID="PanesDropDownList" CssClass="panesDropDown" runat="server"></asp:DropDownList>
    &nbsp;&nbsp;&nbsp;
    <asp:DropDownList ID="ExportTypeDropDown" CssClass="panesDropDown" runat="server">
        <asp:ListItem Text="PDF Adobe Portable Document (*.Pdf)" Selected="True" Value="Pdf"></asp:ListItem>
        <asp:ListItem Text="Mircrosoft Word 97 - 2007 (*.Doc)" Value="Doc"></asp:ListItem>
        <asp:ListItem Text="Office Open XML WordprocessingML (*.Docx)" Value="Docx"></asp:ListItem>
        <asp:ListItem Text="ODF Text Document (*.Odt)" Value="Odt"></asp:ListItem>
        <asp:ListItem Text="Tiff  Image/s (*.Tiff)" Value="Tiff"></asp:ListItem>
        <asp:ListItem Text="JPEG Image (*.Jpeg)" Value="Jpeg"></asp:ListItem>
        <asp:ListItem Text="PNG Image (*.Png)" Value="Png"></asp:ListItem>
    </asp:DropDownList>
    &nbsp;&nbsp;&nbsp;
    <asp:Button ID="ExportButton" OnClientClick="return ButtonClicked();" ResourceKey="ExportButton" runat="server" Text="Export" OnClick="ExportButton_Click" />
   </div>

<div style="clear: both"></div>
