<%@ Page Language="C#" AutoEventWireup="True" CodeBehind="Default.aspx.cs" EnableEventValidation="false"
    Inherits="Aspose.Words.ListViewExport.Website.Default" %>

<%@ Register TagPrefix="Aspose" Namespace="Aspose.Words.ListViewExport" Assembly="Aspose.Words.ListViewExport" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Aspose .NET Export ListView Data to Word</title>
    <style>
        .myClass
        {
            width: 800px;
            text-align: right;
            margin-bottom: 5px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h1>
            Aspose.Words ListViw Export to Word</h1>
        <Aspose:ExportListViewToWord ID="ExportListViewToWord1" GroupPlaceholderID="groupPlaceHolder1"
            ItemPlaceholderID="itemPlaceHolder1" ExportButtonText="Export to Word" ExportButtonCssClass="myClass"
            ExportOutputFormat="Doc" ExportInLandscape="true" ExportOutputPathOnServer="c:\\temp"
            ExportFileHeading="<h4>Example Report</h4>" LicenseFilePath="c:\\inetpub\\Aspose.Words.lic"
            runat="server" CellPadding="4" ExportMaximumRecords="100" OnPagePropertiesChanging="ExportListViewToWord1_PagePropertiesChanging">
            <LayoutTemplate>
                <table cellpadding="0" border="1" width="800px" cellspacing="0">
                    <tr>
                        <th>
                            Product Id
                        </th>
                        <th>
                            Product Name
                        </th>
                        <th>
                            Units In Stock
                        </th>
                    </tr>
                    <asp:PlaceHolder runat="server" ID="groupPlaceHolder1"></asp:PlaceHolder>
                    <tr>
                        <td colspan="3">
                            <asp:DataPager ID="DataPager1" runat="server" PagedControlID="ExportListViewToWord1"
                                PageSize="10">
                                <Fields>
                                    <asp:NextPreviousPagerField ButtonType="Link" ShowFirstPageButton="false" ShowPreviousPageButton="true"
                                        ShowNextPageButton="false" />
                                    <asp:NumericPagerField ButtonType="Link" />
                                    <asp:NextPreviousPagerField ButtonType="Link" ShowNextPageButton="true" ShowLastPageButton="false"
                                        ShowPreviousPageButton="false" />
                                </Fields>
                            </asp:DataPager>
                        </td>
                    </tr>
                </table>
            </LayoutTemplate>
            <GroupTemplate>
                <tr>
                    <asp:PlaceHolder runat="server" ID="itemPlaceHolder1"></asp:PlaceHolder>
                </tr>
            </GroupTemplate>
            <ItemTemplate>
                <td>
                    <%# Eval("Product Id")%>
                </td>
                <td>
                    <%# Eval("Product Name")%>
                </td>
                <td>
                    <%# Eval("Units In Stock")%>
                </td>
            </ItemTemplate>
        </Aspose:ExportListViewToWord>
    </div>
    <br />
    </form>
</body>
</html>
