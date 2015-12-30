<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="View.ascx.cs" Inherits="Aspose.DotNetNuke.Modules.AsposeDNNQuoteGeneratorWord.View" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<link href="/DesktopModules/Aspose.DNN.QuoteGenerator.Word/Css/DNNQuoteGenerator.css"
    rel="stylesheet" type="text/css" media="all" />
<script type="text/javascript" language="javascript">
    function AllowNumericOnly(key) {

        var charCode = (key.which) ? key.which : charCode;
        //alert(charCode);
        if ((charCode >= 48 && charCode <= 57) || (charCode >= 96 && charCode <= 105) || charCode == 8 || charCode == 9 || charCode == 37 || charCode == 39 || charCode == 46)
            return true;

        return false;
    }
    function AllowNumericDecimalsOnly(key) {

        var charCode = (key.which) ? key.which : charCode;
        if ((charCode >= 48 && charCode <= 57) || (charCode >= 96 && charCode <= 105) || charCode == 8 || charCode == 9 || charCode == 110 || charCode == 37 || charCode == 39 || charCode == 46)
            return true;

        return false;
    }
</script>
<div class="aspsoeQuote">
    <h2>
        <img src="/DesktopModules/Aspose.DNN.QuoteGenerator.Word/Images/aspose_logo.gif" />
        Aspose .NET Quote Generator for DNN using Aspose.Words
    </h2>
    <h4>
        <asp:Label ID="lblMessage" runat="server" Font-Bold="true" Text="" ForeColor="Maroon"></asp:Label></h4>
    <table width="100%">
        <tr>
            <td>
                <table width="100%" class="tblBack">
                    <tr>
                        <td colspan="2">
                            <table width="100%">
                                <tr>
                                    <td width="35%">Document Caption (Quote/Invoice/Other)<br />
                                        <asp:TextBox ID="txtDocCaption" CssClass="input" runat="server" Width="50%" Text="QUOTE"
                                            placeholder="Enter document caption (Quote/Invoice/Other)" MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator
                                                CssClass="errors" ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtDocCaption"
                                                ErrorMessage="*Required" ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                    <td width="35%">Document Refference No<br />
                                        <asp:TextBox ID="txtDocNo" CssClass="input" runat="server" Width="50%" Text="" placeholder="Enter refference no"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="RequiredFieldValidator3" runat="server" ControlToValidate="txtDocNo" ErrorMessage="*Required"
                                                ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                    <td width="30%">Date<br />
                                        <asp:TextBox ID="txtDocDate" CssClass="input" runat="server" Width="90%" Text="QUOTE"
                                            placeholder="Enter date" MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator
                                                CssClass="errors" ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtDocDate"
                                                ErrorMessage="*Required" ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td width="48%" class="tdBack">
                            <table width="98%">
                                <tr>
                                    <td colspan="2">
                                        <h3>Sender Information</h3>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Company Logo
                                    </td>
                                    <td>
                                        <asp:FileUpload ID="fuCompanyLogo" CssClass="input" runat="server" />
                                    </td>
                                </tr>
                                <tr>
                                    <td>Company Name
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCompanyName" CssClass="input" runat="server" placeholder="Enter company name"
                                            MaxLength="50"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Company Address
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCompanyAddress" CssClass="input" runat="server" placeholder="Enter company address"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="rfvCompanyAddress" runat="server" ControlToValidate="txtCompanyAddress" ErrorMessage="*Required"
                                                ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <td>&nbsp;
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCompanyStateZip" CssClass="input" runat="server" placeholder="Enter state, zip code"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="rfvCompanyStateZip" runat="server" ControlToValidate="txtCompanyStateZip"
                                                ErrorMessage="*Required" ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <td>&nbsp;
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCompanyCountry" CssClass="input" runat="server" placeholder="Enter country"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="rfvCompanyCountry" runat="server" ControlToValidate="txtCompanyCountry" ErrorMessage="*Required"
                                                ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td width="52%" class="tdBack">
                            <table width="98%">
                                <tr>
                                    <td colspan="2">
                                        <h3>Receiver Information</h3>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Customer Name
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCustomerName" CssClass="input" runat="server" placeholder="Enter customer Name"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="rfvCustomerName" runat="server" ControlToValidate="txtCustomerName" ErrorMessage="*Required"
                                                ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Customer Address
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCustomerAddress" CssClass="input" runat="server" placeholder="Enter customer Address"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="rfvCustomerAddress" runat="server" ControlToValidate="txtCustomerAddress"
                                                ErrorMessage="*Required" ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <td>&nbsp;
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCustomerStateZip" CssClass="input" runat="server" placeholder="Enter state, zip code"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="rfvCustomerStateZip" runat="server" ControlToValidate="txtCustomerStateZip"
                                                ErrorMessage="*Required" ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <td>&nbsp;
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtCustomerCountry" CssClass="input" runat="server" placeholder="Enter country"
                                            MaxLength="50"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator CssClass="errors"
                                                ID="rfvCustomerCountry" runat="server" ControlToValidate="txtCustomerCountry"
                                                ErrorMessage="*Required" ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="vertical-align: baseline !important;">
                            <h3>Products/ Items</h3>
                        </td>
                        <td align="right">
                            <asp:Button ID="btnAddProducts" runat="server" CssClass="buttonClass" Text="Reset Rows To"
                                OnClick="btnAddProducts_Click" />
                            &nbsp;<asp:TextBox ID="txtAddProductRows" CssClass="input" runat="server" Width="25px"
                                MaxLength="2" Text="3" onkeydown=" return AllowNumericOnly(event);"></asp:TextBox>&nbsp;<asp:RequiredFieldValidator
                                    CssClass="errors" ID="rfvAddProductRows" runat="server" ControlToValidate="txtAddProductRows"
                                    ErrorMessage="*Required" ValidationGroup="vgInvoice" Display="Dynamic"></asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <asp:GridView ID="grdInvoiceProducts" TabIndex="14" Width="100%" runat="server" AllowPaging="false"
                                AutoGenerateColumns="False" GridLines="None" DataKeyNames="ProductId" EmptyDataText="There is no record found."
                                ShowHeaderWhenEmpty="True" PageSize="99" BackColor="White" CellPadding="3" ForeColor="Black"
                                ShowFooter="false" OnRowDataBound="grdInvoiceProducts_RowDataBound">
                                <HeaderStyle CssClass="gridViewHeader" Font-Bold="True"></HeaderStyle>
                                <RowStyle CssClass="gridViewRow"></RowStyle>
                                <AlternatingRowStyle CssClass="gridViewAlternateRow"></AlternatingRowStyle>
                                <Columns>
                                    <asp:TemplateField HeaderText="No" ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle"
                                        HeaderStyle-HorizontalAlign="Center" HeaderStyle-VerticalAlign="Middle" HeaderStyle-Width="45px">
                                        <ItemTemplate>
                                            <asp:Label ID="lblNo" Text='<%# Convert.ToInt64(DataBinder.Eval(Container, "RowIndex")) + 1 %>'
                                                runat="server"></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Description" ItemStyle-HorizontalAlign="Left" ItemStyle-VerticalAlign="Middle"
                                        HeaderStyle-HorizontalAlign="Left" HeaderStyle-VerticalAlign="Middle" HeaderStyle-Width="100%">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtProductDescription" CssClass="input" runat="server" Text=""></asp:TextBox>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Price" ItemStyle-HorizontalAlign="Left" ItemStyle-VerticalAlign="Middle"
                                        HeaderStyle-HorizontalAlign="Center" HeaderStyle-VerticalAlign="Middle">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtProductPrice" CssClass="input" Width="150px" runat="server" Text="0.00"
                                                onkeydown=" return AllowNumericDecimalsOnly(event);"></asp:TextBox>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Quantity" ItemStyle-HorizontalAlign="Left" ItemStyle-VerticalAlign="Middle"
                                        HeaderStyle-HorizontalAlign="Center" HeaderStyle-VerticalAlign="Middle">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtProductQuantity" CssClass="input" Width="80px" runat="server"
                                                Text="1" onkeydown=" return AllowNumericOnly(event);"></asp:TextBox>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="VAT" ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle"
                                        HeaderStyle-HorizontalAlign="Center" HeaderStyle-VerticalAlign="Middle">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlProductVAT" Width="80px" CssClass="input" runat="server">
                                            </asp:DropDownList>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <h3>Note / Description</h3>
                        </td>
                        <td>
                            <h3>Terms & Condtions</h3>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:TextBox ID="txtDescription" CssClass="input" ToolTip="Max length is 300 characters"
                                runat="server" placeholder="Enter special note or description" Width="90%" Height="50px"
                                TextMode="MultiLine" Rows="3" MaxLength="300"></asp:TextBox>
                        </td>
                        <td>
                            <asp:TextBox ID="txtTC" CssClass="input" ToolTip="Max length is 300 characters" runat="server"
                                placeholder="Enter terms & condistions if applicable" Width="90%" Height="50px"
                                TextMode="MultiLine" Rows="3" MaxLength="300"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="right">
                            <asp:Button ID="btnClearForm" CssClass="buttonClass" runat="server" Text="Clear Fields"
                                OnClick="btnClearForm_Click" />&nbsp;
                            <asp:DropDownList ID="ExportTypeDropDown" Width="287px" CssClass="input" runat="server">
                                <asp:ListItem Text="PDF Adobe Portable Document (*.Pdf)" Selected="True" Value="Pdf"></asp:ListItem>
                                <asp:ListItem Text="Mircrosoft Word 97 - 2007 (*.Doc)" Value="Doc"></asp:ListItem>
                                <asp:ListItem Text="Office Open XML WordprocessingML (*.Docx)" Value="Docx"></asp:ListItem>
                                <asp:ListItem Text="ODF Text Document (*.Odt)" Value="Odt"></asp:ListItem>
                                <asp:ListItem Text="Tiff  Image/s (*.Tiff)" Value="Tiff"></asp:ListItem>
                                <asp:ListItem Text="JPEG Image (*.Jpeg)" Value="Jpeg"></asp:ListItem>
                                <asp:ListItem Text="PNG Image (*.Png)" Value="Png"></asp:ListItem>
                            </asp:DropDownList>
                            &nbsp
                            <asp:Button ID="btnGeneratePDF" CssClass="buttonClass" runat="server" Text="Export Now"
                                OnClick="btnGeneratePDF_Click" ValidationGroup="vgInvoice" />
                            <p>
                                This is created to demonstrate the <a href="http://www.aspose.com/.net/word-component.aspx"
                                    title="Aspose.Words for .Net (Umbraco)" target="_blank">Aspose.Words for .Net (DNN)</a>.
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</div>
