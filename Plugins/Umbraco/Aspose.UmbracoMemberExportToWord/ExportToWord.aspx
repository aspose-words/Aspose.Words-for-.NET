<%@ Page Language="C#" AutoEventWireup="True" MasterPageFile="~/umbraco/masterpages/umbracoPage.Master"
    CodeBehind="ExportToWord.aspx.cs" Inherits="Aspose.UmbracoMemberExportToWord.AsposeMemberExport" %>

<%@ Register Assembly="controls" Namespace="umbraco.uicontrols" TagPrefix="umbraco" %>
<asp:Content ID="Content2" ContentPlaceHolderID="head" runat="server">
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
    <style type="text/css">
        table.exportToWordGridView
        {
            border: 1px solid #dbdbdb !important;
        }
        
        table.exportToWordGridView td, table.exportToWordGridView th
        {
            padding: 5px;
        }
        
        table.exportToWordGridView thead
        {
            background-color: #f8f8f8;
            border: 1px solid #dbdbdb !important;
        }
    </style>
</asp:Content>
<asp:Content ID="Content1" ContentPlaceHolderID="body" runat="server">
    <umbraco:UmbracoPanel ID="UmbracoPanel" Text="Aspose Member Export to Word 1.0" runat="server">
        <asp:PlaceHolder ID="MemberExportPlaceHolder" runat="server">
            <div class="umb-pane ng-scope">
                <asp:Label ID="ErrorLabel" ForeColor="Red" Visible="false" runat="server"></asp:Label>
                <p>
                    Please select output format and select one or more members to export.
                </p>
                <umbraco:Pane ID="ExportTypePane" runat="server">
                    <umbraco:PropertyPanel ID="ExportTypePropertyPanel" Text="Export Output Format" runat="server">
                        <asp:DropDownList ID="ExportTypeDropDown" runat="server" Width="50%">
                            <asp:ListItem Text="PDF Adobe Portable Document (*.Pdf)" Value="Pdf">
                            </asp:ListItem>
                            <asp:ListItem Text="Mircrosoft Word 97 - 2007 (*.Doc)" Selected="True" Value="Doc">
                            </asp:ListItem>
                            <asp:ListItem Text="Office Open XML WordprocessingML (*.Docx)" Value="Docx">
                            </asp:ListItem>
                            <asp:ListItem Text="ODF Text Document (*.Odt)" Value="Odt">
                            </asp:ListItem>
                            <asp:ListItem Text="Tiff  Image/s (*.Tiff)" Value="Tiff">
                            </asp:ListItem>
                            <asp:ListItem Text="JPEG Image (*.Jpeg)" Value="Jpeg">
                            </asp:ListItem>
                            <asp:ListItem Text="PNG Image (*.Png)" Value="Png">
                            </asp:ListItem>
                        </asp:DropDownList>
                    </umbraco:PropertyPanel>
                        <br />
                        &nbsp;<br />
                        <asp:GridView ID="UmbracoMembersGridView" EmptyDataText="There are no members." Width="100%"
                            EmptyDataRowStyle-CssClass="emptyClass" GridLines="None" BorderWidth="0" AutoGenerateColumns="False"
                            HeaderStyle-CssClass="" EnableViewState="true" CssClass="table table-striped exportToWordGridView"
                            DataKeyNames="Email" ClientIDMode="Static" runat="server">
                            <Columns>
                                <asp:TemplateField HeaderStyle-CssClass="rgHeader" HeaderStyle-Width="35px">
                                    <HeaderTemplate>
                                        <asp:CheckBox ID="SelectAllCheckBox" CssClass="selectAllCheckBox" runat="server" />
                                    </HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="SelectedCheckBox" CssClass="selectableCheckBox" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Text" HeaderStyle-CssClass="rgHeader" HeaderText="Name" />
                                <asp:BoundField DataField="LoginName" HeaderStyle-CssClass="rgHeader" HeaderText="LoginName" />
                                <asp:BoundField DataField="Email" HeaderStyle-CssClass="rgHeader" HeaderText="Email" />
                                <asp:BoundField DataField="CreateDateTime" HeaderStyle-CssClass="rgHeader" HeaderText="CreateDateTime" />
                            </Columns>
                        </asp:GridView>
                </umbraco:Pane>
        </asp:PlaceHolder>
        <asp:PlaceHolder ID="SavePlaceHolder" runat="server">
            <umbraco:Pane ID="ExportPane" runat="server">
                <asp:Button ID="ExportButton" runat="server" Text="Export" OnClick="ExportButton_Click"
                    CssClass="btn btn-success" />
            </umbraco:Pane>
        </asp:PlaceHolder>
        </div>
    </umbraco:UmbracoPanel>
</asp:Content>
