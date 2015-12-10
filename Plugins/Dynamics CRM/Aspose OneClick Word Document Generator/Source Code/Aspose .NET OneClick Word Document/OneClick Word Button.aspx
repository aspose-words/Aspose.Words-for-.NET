<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="OneClick Word Button.aspx.cs" Inherits="Aspose.NET_OneClick_Word_Document.OneClick_Word_Button" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="AsposeStyles.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
        <asp:Label ID="LBL_Message" runat="server" Text="" ForeColor="Red" Visible="false"></asp:Label>
        <asp:Button ID="BTN_Download" runat="server" Text="Download" CssClass="topbutton" Visible="false" OnClick="BTN_Download_Click" />&nbsp;&nbsp;&nbsp;<asp:Button ID="BTN_AttachToNote" runat="server" Text="Attach To Note" CssClass="topbutton" Visible="false" OnClick="BTN_AttachToNote_Click" />
    </form>
</body>
</html>
