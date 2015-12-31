<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="OneClick Word Page.aspx.cs" Inherits="Aspose.NET_OneClick_Word_Document.OneClick_Word_Page" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>OneClick Word Document Generator Page</title>
    <link href="AsposeStyles.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="Main">
            <h2>Aspose .NET OneClick Word Document Generator
            </h2>
            <div class="Label">Select Template and Generate Quote</div>
            <br />
            <asp:Label ID="LBL_Message" runat="server" Text="" ForeColor="Red"></asp:Label>
            <div class="Width100">
                <table>

                    <tr>
                        <th class="Label Width30">Template
                        </th>
                        <td class="Value">
                            <asp:DropDownList ID="DDL_Templates" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <th class="Label Width30">File Format
                        </th>
                        <td class="Value">
                            <asp:DropDownList ID="DDL_FileFormat" runat="server">
                                <asp:ListItem Text="Docx"></asp:ListItem>
                                <asp:ListItem Text="Doc"></asp:ListItem>
                                <asp:ListItem Text="rtf"></asp:ListItem>
                                <asp:ListItem Text="bmp"></asp:ListItem>
                                <asp:ListItem Text="html"></asp:ListItem>
                                <asp:ListItem Text="jpeg"></asp:ListItem>
                                <asp:ListItem Text="pdf"></asp:ListItem>
                                <asp:ListItem Text="png"></asp:ListItem>
                                <asp:ListItem Text="text"></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <th class="Label Width30">Action
                        </th>
                        <td class="Value">
                            <asp:DropDownList ID="DDL_Action" runat="server">
                                <asp:ListItem Text="Download"></asp:ListItem>
                                <asp:ListItem Text="Attach to Note"></asp:ListItem>
                                <%--<asp:ListItem Text="Attach to Email"></asp:ListItem>--%>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <th class="Label Width30">Generated File Name
                        </th>
                        <td class="Value">
                            <asp:TextBox ID="TXT_FileName" runat="server"></asp:TextBox>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        <div class="footer">
            <div class="buttons-left">
            </div>
            <div class="buttons-right">
                <asp:Button ID="BTN_Generate" runat="server" Text="Generate" CssClass="footerbutton" OnClick="BTN_Generate_Click" />
                <button class="footerbutton" type="button" onclick="window.close();">
                    Cancel
                </button>
            </div>
        </div>
    </form>
</body>
</html>
