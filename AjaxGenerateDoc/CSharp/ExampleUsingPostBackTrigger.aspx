<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExampleUsingPostBackTrigger.aspx.cs" Inherits="AjaxGenerateDoc.ExampleUsingPostBackTrigger" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Aspose.Words with AJAX - Example 3</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <div>
            <span>
                This example shows how to invoke Aspose.Words for generating a document with data from a GridView control.<br />
                In this example PostBack trigger is used (full PostBack is invoked).
                <br />
            </span>
            <br />
            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                <ContentTemplate>
                    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" OnRowCommand="GridView1_RowCommand">
                        <Columns>
                            <asp:ButtonField ButtonType="Link" Text="Generate" HeaderText="Generate" CommandName="generate" />
                            <asp:BoundField DataField="Name" HeaderText="Name"  />
                            <asp:BoundField DataField="Company" HeaderText="Company" />
                        </Columns>
                    </asp:GridView>
                </ContentTemplate>
                <Triggers>
                    <asp:PostBackTrigger ControlID="GridView1" />
                </Triggers>
            </asp:UpdatePanel>
        
        </div>
    </form>
</body>
</html>
