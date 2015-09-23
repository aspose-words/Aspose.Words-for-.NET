<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="GenerateQuote.aspx.cs" Inherits="Aspose.QuoteGenerator.GenerateQuote" ValidateRequest="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="http://cdn.ckeditor.com/4.5.3/full/ckeditor.js"></script>
    <link href="AsposeStyles.css" rel="stylesheet" />
    <script src="../ClientGlobalContext.js.aspx" type="text/javascript"></script>
    <script src="JQuery.js"></script>

</head>
<body>
    <form id="form1" runat="server">
        <div id="msgDiv" runat="server" class="loaderStyle" style="display: none;">
            <img alt='' src='/_imgs/AdvFind/progress.gif' /><br />
            Loading...</div>
        <div class="Main" style="height: 10% !important;">
            <h2>Generate Document
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
                            <asp:DropDownList ID="DDL_Templates" runat="server" OnSelectedIndexChanged="DDL_Templates_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        <textarea name="editor1" id="editor1" runat="server"></textarea>
        <div class="footer">
            <div class="buttons-left">
            </div>
            <div class="buttons-right">
                <asp:Button ID="BTN_Download" runat="server" Text="Download Document" CssClass="footerbutton" OnClick="BTN_Download_Click" Visible="false" />
                <asp:Button ID="BTN_Attach" runat="server" Text="Attach to Quote" CssClass="footerbutton" OnClick="BTN_Attach_Click" />
                <button class="footerbutton" type="button" onclick="window.close();">
                    Cancel
                </button>
            </div>
        </div>
        <script>
            CKEDITOR.replace("editor1");
        </script>
    </form>
</body>
</html>
