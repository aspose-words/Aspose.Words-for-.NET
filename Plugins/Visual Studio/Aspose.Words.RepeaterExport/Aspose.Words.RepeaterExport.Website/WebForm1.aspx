<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="Aspose.Words.RepeaterExport.Website.WebForm1" %>
<%@ Register TagPrefix="aspose" Namespace="Aspose.Words.RepeaterExport" Assembly="Aspose.Words.RepeaterExport" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <aspose:ExportRepeaterToWord ID="ExportRepeaterToWord1" ExportButtonText="Export to Word"
             ExportOutputFormat="Doc"
            ExportInLandscape="true" ExportOutputPathOnServer="c:\\temp" ExportFileHeading="<h4>Example Report</h4>"
                        LicenseFilePath="e:\\Aspose\\Aspose.Words.lic"
            runat="server" >
          <HeaderTemplate>
             <table border="1">
          </HeaderTemplate>

          <ItemTemplate>
             <tr>
                <td> <%# Container.DataItem %> </td>
             </tr>
          </ItemTemplate>

          <FooterTemplate>
             </table>
          </FooterTemplate>
        </aspose:ExportRepeaterToWord>
    </div>
    </form>
</body>
</html>
