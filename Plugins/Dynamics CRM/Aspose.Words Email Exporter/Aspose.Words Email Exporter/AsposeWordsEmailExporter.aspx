<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AsposeWordsEmailExporter.aspx.cs" Inherits="Aspose.Words_Email_Exporter.AsposeWordsEmailExporter" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Aspose.Words Email Exporter</title>
    <link href="AsposeStyles.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="Main">
            <h2>Aspose.Words Email Exporter
            </h2>
            <div class="Label">Select Export Options below</div>
            <br />
            <asp:Label ID="LBL_Message" runat="server" Text="" ForeColor="Red"></asp:Label>
            <div class="Width100">
                <table>
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
                                <asp:ListItem Text="Attach to This Email"></asp:ListItem>
                                <%--<asp:ListItem Text="Create New Email"></asp:ListItem>
                                <asp:ListItem Text="Attach to Record"></asp:ListItem>--%>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <th class="Label Width30">File Name
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
