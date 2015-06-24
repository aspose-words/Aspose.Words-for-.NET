<%@ Page Language="C#" AutoEventWireup="True" CodeBehind="Default.aspx.cs" Inherits="Aspose.Words.GridViewExport.Website.Default" %>

<%@ Register TagPrefix="aspose" Namespace="Aspose.Words.GridViewExport" Assembly="Aspose.Words.GridViewExport" %>
<!DOCTYPE html>
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
        <aspose:ExportGridViewToWord Width="800px" ID="ExportGridViewToWord1" ExportButtonText="Export to Word"
            CssClass="table table-hover table-bordered" ExportButtonCssClass="myClass" ExportOutputFormat="Doc"
            ExportInLandscape="true" ExportOutputPathOnServer="c:\\temp" ExportFileHeading="<h4>Example Report</h4>"
            OnPageIndexChanging="ExportGridViewToWord1_PageIndexChanging" PageSize="5" AllowPaging="True"
            LicenseFilePath="c:\\inetpub\\Aspose.Words.lic"
            runat="server" CellPadding="4" ForeColor="#333333" GridLines="Both">
        </aspose:ExportGridViewToWord>
    </div>
    </form>
</body>
</html>
