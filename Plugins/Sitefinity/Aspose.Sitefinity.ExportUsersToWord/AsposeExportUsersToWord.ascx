<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="AsposeExportUsersToWord.ascx.cs" Inherits="Aspose.Sitefinity.ExportUsersToWord.AsposeExportUsersToWord" %>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.1/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" media="all" href="<%= ResolveUrl("~/Addons/Aspose.SiteFinity.ExportUsersToExcel/css/style.css") %>" />
<style type="text/css">
    .error{
        padding: 15px;
    }
</style>
<script type="text/javascript" language="javascript">

    $(document).ready(function () {
        $('.selectAllCheckBox input[type="checkbox"]').click(function (event) {  //on click
            if (this.checked) { // check select status
                $('.selectableCheckBox input[type="checkbox"]').each(function () { //loop through each checkbox
                    this.checked = true;  //select all checkboxes with class "checkbox1"              
                });
            } else {
                $('.selectableCheckBox input[type="checkbox"]').each(function () { //loop through each checkbox
                    this.checked = false; //deselect all checkboxes with class "checkbox1"                      
                });
            }
        });
    });
</script>
<div style="margin-top: 100px;" class="container">
    <h2 class="sub-header">Export Sitefinity Users to Word / Pdf</h2>
    <p id="NoRowSelectedErrorDiv" runat="server" visible="false" class="bg-danger error">Please select one or more users to export.</p>
    <div class="row">
        <ul class="list-inline pull-right">
            <li>
                <asp:DropDownList ID="ExportTypeDropDown" runat="server" CssClass="pull-right form-control">
                    <asp:ListItem Text="Portable Document Format (*.pdf)" Value="pdf"></asp:ListItem>
                    <asp:ListItem Text="Word Document (*.docx)" Selected="True" Value="docx"></asp:ListItem>    
                    <asp:ListItem Text="Word 97-2003 Document (*.doc)" Value="doc"></asp:ListItem>
                    <asp:ListItem Text="Word 97-2003 Document Template (*.dot)" Value="dot"></asp:ListItem>
                    <asp:ListItem Text="Word Document Template (*.dotx)" Value="dotx"></asp:ListItem>
                    <asp:ListItem Text="Word Open XML Macro - Enabled Document (*.docm)" Value="docm"></asp:ListItem>
                    <asp:ListItem Text="Word Open XML Macro - Enabled Template (*.dotm)" Value="dotm"></asp:ListItem>
                    <asp:ListItem Text="OpenDocument Format (*.odt)" Value="odt"></asp:ListItem>
                    <asp:ListItem Text="Opent Office Document Format (*.ott)" Value="ott"></asp:ListItem>
                    <asp:ListItem Text="Rich Text Format (*.rtf)" Value="rtf"></asp:ListItem>
                    <asp:ListItem Text="Text (Tab delimited) (*.txt)" Value="txt"></asp:ListItem>
                </asp:DropDownList>
            </li>
            <li>
                <asp:Button ID="ExportButton" CssClass="btn btn-primary pull-right" runat="server" Text="Export" OnClick="ExportButton_Click"></asp:Button>
            </li>
        </ul>
    </div>
    <div class="row">
        <asp:GridView ID="SitefinityUsersGridView" EmptyDataText="There are no users." Width="100%" EmptyDataRowStyle-CssClass="emptyClass"
            GridLines="None" BorderWidth="0" AutoGenerateColumns="false"
            CssClass="table table-striped" DataKeyNames="Email" ClientIDMode="Static" runat="server">
            <Columns>
                <asp:TemplateField HeaderStyle-CssClass="rgHeader" HeaderStyle-Width="20px">
                    <HeaderTemplate>
                        <asp:CheckBox ID="SelectAllCheckBox" CssClass="selectAllCheckBox" runat="server" />
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="SelectedCheckBox" CssClass="selectableCheckBox" runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Username" HeaderText="Username" />
                <asp:BoundField DataField="FirstName" HeaderText="Firstname" />
                <asp:BoundField DataField="LastName" HeaderText="Lastname" />
                <asp:BoundField DataField="Email" HeaderText="Email" />
            </Columns>
        </asp:GridView>
    </div>
</div>