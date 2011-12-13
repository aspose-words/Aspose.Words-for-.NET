<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebRole._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <strong>Simple Upload and Convert</strong><br />
    
        Step 1. Select file to upload:
        <asp:FileUpload ID="SrcFileUpload" runat="server" />
&nbsp;(demo limit 512Kb)<br />
        Step 2. Select output format:         <asp:DropDownList ID="DstFormatDropDownList" runat="server">
            <asp:ListItem Selected="True" Value="PDF">PDF - Adobe Portable Document Format</asp:ListItem>
            <asp:ListItem Value="XPS">XPS - Microsoft XML Paper Specification</asp:ListItem>
            <asp:ListItem Value="DOC">DOC - Microsoft Word 97-2007 Document</asp:ListItem>
            <asp:ListItem Value="DOCX">DOCX - Office Open XML</asp:ListItem>
            <asp:ListItem Value="RTF">RTF - Rich Text Format</asp:ListItem>
            <asp:ListItem Value="XML">XML - Microsoft Word 2003 WordprocessingML</asp:ListItem>
            <asp:ListItem Value="ODT">ODT - OpenDocument Text</asp:ListItem>
            <asp:ListItem Value="MHTML">MHTML - Web Page Archive</asp:ListItem>
            <asp:ListItem Value="EPUB">EPUB - IDPF eBook</asp:ListItem>
            <asp:ListItem Value="TXT">TXT - Plain Text</asp:ListItem>
        </asp:DropDownList>
        <br />
        Step 3.
        <asp:Button ID="SubmitButton" runat="server" onclick="SubmitButton_Click" 
            Text="Submit" />
    </div>
    </form>
</body>
</html>
