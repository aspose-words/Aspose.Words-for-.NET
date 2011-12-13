<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebRole._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        1. Enter some test data, e.g. whom this document is for:
        <asp:TextBox ID="NameTextBox" runat="server">James Bond</asp:TextBox>
        <br />
        2. Select the format for the document:         <asp:DropDownList ID="DstFormatDropDownList" runat="server">
            <asp:ListItem Value="DOC">DOC - Microsoft Word 97-2007 Document</asp:ListItem>
            <asp:ListItem Value="DOCX">DOCX - Office Open XML</asp:ListItem>
            <asp:ListItem Value="PDF">PDF - Adobe Portable Document Format</asp:ListItem>
            <asp:ListItem Value="XPS">XPS - Microsoft XML Paper Specification</asp:ListItem>
            <asp:ListItem Value="ODT">ODT - OpenDocument Text Format</asp:ListItem>
            <asp:ListItem Value="RTF">RTF - Rich Text Format</asp:ListItem>
            <asp:ListItem Value="XML">XML - Microsoft Word 2003 WordprocessingML</asp:ListItem>
            <asp:ListItem Value="MHTML">MHTML - Web Page Archive</asp:ListItem>
        </asp:DropDownList>
        <br />
        3. Click this button to generate a document and save to the Windows Azure 
        Storage Blob Service:         <asp:Button ID="GenerateButton" runat="server" onclick="GenerateButton_Click" 
            Text="Generate" />
    
    </div>
    <hr />
    <div>
        Documents in the Windows Azure Storage:<br />
        <asp:GridView ID="BlobGridView" runat="server" CellPadding="4" 
            EnableModelValidation="True" ForeColor="#333333" GridLines="None">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="Uri" Text="Download" 
                    Target="_blank" />
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        </asp:GridView>
        <br />
        <asp:Button ID="ClearStorageButton" runat="server" 
            onclick="ClearStorageButton_Click" Text="Clear Storage" />
    </div>
    </form>
</body>
</html>
