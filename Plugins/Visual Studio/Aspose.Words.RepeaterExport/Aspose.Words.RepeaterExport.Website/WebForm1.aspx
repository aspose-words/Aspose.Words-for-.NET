<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="Aspose.Words.RepeaterExport.Website.WebForm1" %>

<%@ Register TagPrefix="aspose" Namespace="Aspose.Words.RepeaterExport" Assembly="Aspose.Words.RepeaterExport" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .myClass
        {
            clear: both;
        }
        
        .myClass input
        {
            float: right;
        }
    </style>
    <!-- Latest compiled and minified CSS -->
    <link rel="stylesheet" href="http://netdna.bootstrapcdn.com/bootstrap/3.1.1/css/bootstrap.min.css" />
    <!-- Optional theme -->
    <link rel="stylesheet" href="http://netdna.bootstrapcdn.com/bootstrap/3.1.1/css/bootstrap-theme.min.css" />
    <!-- Latest compiled and minified JavaScript -->
    <script type="text/javascript" src="http://netdna.bootstrapcdn.com/bootstrap/3.1.1/js/bootstrap.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div style="width: 800px">
            <aspose:ExportRepeaterToWord ID="ExportRepeaterToWord1" ExportButtonText="Export to Word"
                ExportButtonCssClass="myClass" ExportOutputFormat="Doc" ExportInLandscape="true"
                ExportOutputPathOnServer="E:\\temp" ExportFileHeading="<h4>Example Report</h4>"
                LicenseFilePath="E:\\Aspose\\Aspose.Total.lic" runat="server">
                <HeaderTemplate>
                    <table class="table table-hover table-bordered">
                        <tr>
                            <th>
                                Product ID
                            </th>
                            <th>
                                Product Name
                            </th>
                            <th>
                                Units In Stock
                            </th>
                        </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td>
                            <%# Eval("Product ID") %>
                        </td>
                        <td>
                            <%# Eval("Product Name")%>
                        </td>
                        <td>
                            <%# Eval("Units In Stock")%>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </table>
                </FooterTemplate>
            </aspose:ExportRepeaterToWord>
        </div>
    </div>
    </form>
</body>
</html>
