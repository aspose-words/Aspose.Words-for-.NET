<%@ Page Title="" Language="C#" MasterPageFile="~/site.Master" AutoEventWireup="true" CodeBehind="errorpage.aspx.cs" Inherits="Aspose.App.UI.errorpage" %>


<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
     
   <div class="error404">
    <div class="content404">
    <h4><% = Resources["PageNotFound"] %></h4>
    <h2><% = Resources["404"] %></h2>
    <p><% = Resources["BrokenLinkTitle"] %></p>
    <span><a href="javascript:history.go(-1)" class="backButton"><% = Resources["btnBackHome"] %></a> <a href="https://www.aspose.app/"><% = Resources["HomePage"] %></a> </span>
</div></div>
	        
</asp:Content>

