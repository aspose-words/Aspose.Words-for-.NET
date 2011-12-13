<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="ExampleUsingIFrame2.aspx.vb" Inherits="AjaxGenerateDoc.ExampleUsingIFrame2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>Aspose.Words with AJAX - Example 2</title>
</head>
<body>
	<form id="form1" runat="server">
		<asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" />
		<br />
			<span>
				This example shows how to invoke Aspose.Words for generating a document with data from a GridView control.<br />
				In this example IFrame is used. 
				<br />
			</span>
			<br />
			<asp:UpdatePanel ID="UpdatePanel1" runat="server">
				<ContentTemplate>
					<asp:GridView ID="GridView1" runat="server" OnRowDataBound="GridView1_RowDataBound" AutoGenerateColumns="false">
						<Columns>
							<asp:HyperLinkField Text="Generate" HeaderText="Generate" NavigateUrl="#" />
							<asp:BoundField DataField="Name" HeaderText="Name"  />
							<asp:BoundField DataField="Company" HeaderText="Company" />
						</Columns>
					</asp:GridView>                
				</ContentTemplate>
			</asp:UpdatePanel>
	</form>
</body>
</html>
