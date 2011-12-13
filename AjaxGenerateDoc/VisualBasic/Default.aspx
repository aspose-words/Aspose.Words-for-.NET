<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="AjaxGenerateDoc._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
	<title>Aspose.Words with AJAX</title>
</head>
<body>
	<form id="form1" runat="server">
	<div>
		<h1>Using Aspose.Words with AJAX</h1>
		The examples below invoke Aspose.Words to generate a Microsoft Word document when
		a user clicks a button or a link on a web page. The request to generate and the
		produced document are delivered using AJAX functionality in ASP.NET.<br />
		<br />
		<asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/ExampleUsingIFrame1.aspx">Example 1</asp:HyperLink><span> - Displays a progress message while generating a file. An IFrame is used.</span>
		<br />
		<asp:HyperLink ID="HyperLink2" runat="server" NavigateUrl="~/ExampleUsingIFrame2.aspx">Example 2</asp:HyperLink><span> - Uses data from a GridView control to generate a customized document. An IFrame is used.</span>
		<br />
		<asp:HyperLink ID="HyperLink3" runat="server" NavigateUrl="~/ExampleUsingPostBackTrigger.aspx">Example 3</asp:HyperLink><span> - Uses data from a GridView control to generate a customized document, but employs a PostBack trigger
			(full PostBack is invoked).</span></div>
	</form>
</body>
</html>
