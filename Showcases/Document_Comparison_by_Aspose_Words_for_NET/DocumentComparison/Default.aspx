<%@ Page Title="Document Comparison by Aspose.Words for .NET" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="DocumentComparison.Default" %>

<asp:Content ID="HeadContent" ContentPlaceHolderID="head" runat="server">
</asp:Content>


<asp:Content ID="BodyContent" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <input id="txtSessionID" type="hidden" value="<%= Session.SessionID %>" />

    <%--<h1>My Documents</h1>--%>
    <p class="lead margin-top-10">Manage, view and compare Microsoft Word documents.</p>

    <ul id="myTab" class="nav navbar-default nav-tabs margin-bottom-5">
        <li class="active">
            <a href="#MyDocumentsTab" data-toggle="tab">My Documents</a>
        </li>
        <li><a href="#URLsTab" data-toggle="tab">URLs</a></li>
    </ul>



    <asp:Label runat="server" ID="lblCurrentFolder" Text="" CssClass="hidden lblCurrentFolder"></asp:Label>

    <div id="divPageAlert" class="alert hidden">this is alert</div>

    <!-- Tab Contents -->
    <div id="myTabContent" class="tab-content">
        <!-- My Documents tab -->
        <div class="tab-pane fade in active" id="MyDocumentsTab">
            <div class="panel panel-default">
                <!-- Default panel contents -->
                <div class="panel-heading clearfix">
                    <div class="pull-left">
                        <input type="checkbox" id="chSelectAll" />
                        &nbsp;&nbsp;
                        <button class="btn btn-primary" value="" onclick="btnCompare_onClick(); return false;">
                            Compare Documents
                        </button>
                    </div>
                    &nbsp;
            <div class="pull-right">
                <asp:FileUpload CssClass="hidden" ID="FileUpload1" runat="server" onChange="this.form.submit();" />
                <asp:Button CssClass="hidden" ID="Button1" runat="server" Text="Button" />
                <%--<asp:LinkButton ID="btnUpload" runat="server" Text="Upload" OnClick="UploadFile">
                    <span class="glyphicon glyphicon-upload upload-document" aria-hidden="true"></span>
                </asp:LinkButton>--%>
                <button onclick="document.getElementById('<%=FileUpload1.ClientID%>').click(); return false;">
                    Upload Document
                    <span class="glyphicon glyphicon-upload upload-document" aria-hidden="true"></span>
                </button>
                &nbsp;
                <button type="button" data-toggle="modal" data-target="#CreateFolderDialog">
                    Create Folder
                    <span class="glyphicon glyphicon-plus create-folder" aria-hidden="true"></span>
                </button>
            </div>
                </div>
                <asp:ListView ID="GridView1" runat="server" OnItemDataBound="GridView1_ItemDataBound" OnItemCommand="GridView1_ItemCommand">
                    <LayoutTemplate>
                        <table class="table table-bordered table-hover" runat="server" id="DocumentsTable">
                            <tr runat="server" id="itemPlaceholder"></tr>
                        </table>
                    </LayoutTemplate>
                    <ItemTemplate>
                        <tr runat="server">
                            <td>&nbsp;
                        <asp:CheckBox CssClass="select-document" runat="server" />
                                <input type="hidden" class="link-document" value="<%# Eval("FullName")  %>" />
                            </td>
                            <td runat="server">
                                <asp:LinkButton runat="server" ID="lbFolderItem" CommandName="OpenFolder" CssClass=""
                                    OnClientClick='<%# String.Format("return viewDocument(\"{0}\" , \"{1}\")", Eval("FullName").ToString().Replace("\\", "\\\\"), Eval("IsFolder")) %>'
                                    CommandArgument='<%# Eval("Name") + DocumentComparison.Common.separator[0] + Eval("IsFolder") %>'>
                            <%# Eval("Name") %>
                                </asp:LinkButton>
                            </td>
                            <td>
                                <%# DocumentComparison.Common.DisplaySize((long?) Eval("Size")) %>
                            </td>
                            <td>
                                <%# DocumentComparison.Common.FormatDate((DateTime) Eval("LastWriteTime")) %>
                            </td>
                            <td runat="server">
                                <asp:LinkButton ID="lnkDownload" Text="Download" CommandName="DownloadFile"
                                    CommandArgument='<%# Eval("FullName") + DocumentComparison.Common.separator[0] + Eval("IsFolder") %>' runat="server">
                            Download
                        <span class="glyphicon glyphicon-download download-document" aria-hidden="true"></span>
                                </asp:LinkButton>
                            </td>
                            <td runat="server">
                                <asp:LinkButton ID="lnkDelete" Text="Delete" CommandName="DeleteFile" CssClass="delete-confirm"
                                    CommandArgument='<%# Eval("FullName") + DocumentComparison.Common.separator[0] + Eval("IsFolder") %>' runat="server">
                        Delete
                            <span class="glyphicon glyphicon-remove-circle delete-document" aria-hidden="true"></span>
                                </asp:LinkButton>
                            </td>
                        </tr>
                    </ItemTemplate>

                </asp:ListView>
            </div>
        </div>

        <!-- URLs tab -->
        <div class="tab-pane fade" id="URLsTab">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <button class="btn btn-primary" value="" onclick="btnCompareURLs_onClick(); return false;">
                        Compare Documents
                    </button>
                </div>
                <div class="panel-body">
                    
                    <div class="form-group">
                        <label for="txtFirstURL" class="col-sm-2 control-label">First URL</label>
                        <div class="col-sm-10 margin-bottom-10">
                            <input type="text" class="form-control form-inline" id="txtFirstURL"
                                placeholder="Enter First URL" value="http://www.aspose.com/blogs/wp-content/uploads/2010/04/Newsletter-1.docx">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtSecondURL" class="col-sm-2 control-label">Second URL</label>
                        <div class="col-sm-10">
                            <input type="text" class="form-control" id="txtSecondURL"
                                placeholder="Enter Second URL" value="http://www.aspose.com/blogs/wp-content/uploads/2010/04/Newsletter-2.docx">
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    

    <!-- Dialog boxes -->
    <div class="modal fade" id="CreateFolderDialog" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content" id="CreateFolderModalContent">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                    <h4 class="modal-title" id="myModalLabel">Create New Folder</h4>
                </div>
                <div class="modal-body">
                    <div id="divAlert" class="alert hidden">this is alert</div>
                    <div class="input-group">
                        <input id="txtCreateFolder" type="text" class="form-control" placeholder="Enter folder name...">
                        <span class="input-group-btn">
                            <button runat="server" class="btn btn-primary" type="button" onclick="btnCreateFolder_onClick();">Create</button>
                        </span>
                    </div>
                    <!-- /input-group -->
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="PageGeneralDialog" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content" id="PageGeneralModalContent">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                    <h4 class="modal-title" id="PageGeneralLabel"></h4>
                </div>
                <div class="modal-body">
                    <div id="PageGeneralDivAlert" class="alert hidden">this is alert</div>
                </div>
            </div>
        </div>
    </div>

    <!-- Document Viewer -->
    <div class="modal fade" id="DocumentViewerDialog" tabindex="-1" role="dialog" aria-labelledby="DocumentViewerDialogTitle" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content" id="DocumentViewerDialogContent">
                <div class="modal-header">

                    <nav id="DocumentViewerPagination" class="navbar navbar-default" role="navigation">
                        <div class="navbar-header">
                            <span id="DocumentViewerDialogTitle" class="navbar-brand">TutorialsPoint</span>
                        </div>
                        <div>
                            <!--Left Align-->
                            <ul id="DocumentViewerPaginationUL" class="  nav navbar-nav navbar-left">
                                <li>
                                    <a href="#" aria-label="Previous">
                                        <span aria-hidden="true">Pages:</span>
                                    </a>
                                </li>


                                <li class="DocumentViewerPaginationLI"><a href="#">1</a></li>
                            </ul>
                            <p class="navbar-text navbar-right">&nbsp;&nbsp;</p>
                            <button type="button" class="navbar-text navbar-right" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                        </div>
                    </nav>

                    <div id="DocumentViewerSummary">
                        Added <span id="DocumentViewerSummaryAdded" class="label label-primary">4</span> ,
                        Deleted <span id="DocumentViewerSummaryDeleted" class="label label-danger">2</span>
                    </div>

                </div>

                <div class="modal-body">
                    <div id="DocumentViewerAlert" class="alert hidden">this is alert</div>
                    <img class="img-responsive center-block" src="http://localhost:50465/Temp/temp.png" id="CurrentDocumentPage" />
                </div>
            </div>
        </div>
    </div>

    <script src="Default.js"></script>
</asp:Content>
