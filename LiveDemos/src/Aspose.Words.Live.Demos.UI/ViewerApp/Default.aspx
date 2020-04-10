<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Aspose.Words.Live.Demos.UI.Viewer.Default" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width" />
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Cache-Control" content="no-cache">
    <meta http-equiv="Expires" content="Sat, 01 Dec 2001 00:00:00 GMT">

    <title>Free File Format Apps - Aspose Document Viewer</title>
    <link href="https://products.aspose.com/templates/aspose/favicon.ico" rel="shortcut icon" type="image/vnd.microsoft.icon" />
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular-sanitize.min.js"></script>
    
    <link type="text/css" rel="stylesheet" href="/ViewerApp/resources/common/css/all.min.css">
    <link type="text/css" rel="stylesheet" href="/ViewerApp/resources/common/css/v4-shims.min.css">
    <link type="text/css" rel="stylesheet" href="/ViewerApp/resources/common/css/swiper.min.css">
    <link type="text/css" rel="stylesheet" href="/ViewerApp/resources/common/css/circle-progress.css" />
    <link type="text/css" rel="stylesheet" href="/ViewerApp/resources/viewer/css/viewer.css" />
    <link type="text/css" rel="stylesheet" href="/ViewerApp/resources/viewer/css/viewer.mobile.css" />
    <link type="text/css" rel="stylesheet" href="/ViewerApp/resources/viewer/css/viewer-light.css" />
    <link type="text/css" rel="stylesheet" href="https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css" />
    <script type="text/javascript" src="/ViewerApp/resources/common/js/jquery.min.js"></script>
    <script type="text/javascript" src="/ViewerApp/resources/common/js/swiper.min.js"></script>
    <script type="text/javascript" src="/ViewerApp/resources/common/js/es6-promise.auto.js"></script>
    <script type="text/javascript" src="/ViewerApp/resources/viewer/js/viewer.js"></script>
    
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    
    <link rel="stylesheet" href="~/css/editor.css" />

    <script src="/ViewerApp/app.js"></script>
    <script src="/ViewerApp/app.controller.main.js"></script>
    
  
</head>
<body ng-app="AsposeViewerApp" style="padding-top: 70px;">
    <div ng-controller="ViewerController">
        <div class="navbar navbar-inverse navbar-fixed-top" style="margin: 0; background-color: #131313!important; max-height: 50px;">
            <div class="container-fluid">
                <div class="navbar-header">
                    <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target=".navbar-collapse" aria-expanded="false">
                        <span class="sr-only">Toggle navigation</span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                    </button>
                    <a class="navbar-brand" href="<%= CallbackURL %>" style="padding: 5px 15px;">
                        <img src="/ViewerApp/Resources/images/aspose-logo.jpg" alt="Aspose Document Editor App" />
                    </a>
                </div>
                <div class="hidden-xs">
                    <h3 class="navbar-text" style="margin-top: 15px;">
                        <%= ProductName %>
                    </h3>
                    <p class="navbar-text navbar-center" style="margin-top: 18px;">
                        <%= FileName %>
                    </p>

                    <button type="button" class="btn navbar-btn navbar-right closebutton" data-toggle="modal" data-target="#returnModal">
                        <i class="glyphicon glyphicon-remove" style="color: #9d9d9d;"></i>
                    </button>
                    <ul class="nav navbar-nav navbar-right">
                        <li class="dropdown">
                            <button type="button" class="btn btn-success dropdown-toggle" ng-click="Download()" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false"
                                style="margin-top: 7px;">
                                Download&nbsp;<span class="caret">&nbsp;</span>
                            </button>
                            <ul class="dropdown-menu">
                                <li><a href="#" ng-click="Download('.docx')">as DOCX</a></li>
                                <li><a href="#" ng-click="Download('.pdf')">as PDF</a></li>
                                <li><a href="#" ng-click="Download('.html')">as HTML</a></li>
                            </ul>
                        </li>
                    </ul>
                </div>
                <div class="visible-xs-block">
                <div class="collapse navbar-collapse navbar-inverse navbar-left">
                    <ul class="nav navbar-nav pull-right">
                        <li><a href="#" ng-click="Download()" style="color: white;">Download</a></li>
                        <li><a href="#" data-toggle="modal" data-target="#returnModal"style="color: white;">Exit</a></li>
                    </ul>
                </div>
                </div>
            </div>
        </div>

        <div id="alert" class="alert alert-danger" role="alert" style="display: none; z-index: 1050; position: relative; margin-top: 50px;">
            <button type="button" class="close" aria-label="Close" onclick="$('#alert').hide()"><span aria-hidden="true">&times;</span></button>
            <p></p>
        </div>

        <div id="page-loading">
            <img id="htmlloader" src="/ViewerApp/Resources/images/loader.gif" />
            <div id="loader" style="display: none;"></div>
        </div>
        
        <div id="viewer"></div>

        <div id="returnModal" class="modal fade" tabindex="-1" role="dialog">
            <div class="modal-dialog" role="document">
                <div class="modal-content">
                    <div class="modal-body">
                        <p><%= Resources["EditorReturnQuestion"] %></p>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-primary" onclick="closeWindow()">Yes</button>
                        <button type="button" class="btn btn-default" data-dismiss="modal">No</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        'use strict';
        window.asposeViewerAPI = '<%= AsposeViewerApp %>';
        window.fileName = '<%= FileName %>';
        window.productName = '<%= Product %>';
        window.folderName = '<%= FolderName %>';
        window.fileDownloadLink = '<%= FileDownloadLink %>';

        function closeWindow() {
            if (window.parent && window.parent.closeIframe) {
                window.history.back();
                window.parent.closeIframe();
            } else {
                window.location.href = '<%= CallbackURL %>';
            }
        }
    </script>
</body>
</html>
 
