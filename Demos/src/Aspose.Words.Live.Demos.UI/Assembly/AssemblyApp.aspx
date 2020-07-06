<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="AssemblyApp.aspx.cs" Inherits="Aspose.Words.Live.Demos.UI.AssemblyApp" %>


<asp:Content ID="AssemblyContent" ContentPlaceHolderID="MainContent" runat="server">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular-route.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular-animate.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular-aria.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular-resource.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.7.8/angular-messages.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/angular-material/1.1.12/angular-material.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/danialfarid-angular-file-upload/12.2.13/ng-file-upload-all.min.js"></script>

    <script src="/assembly/app.js"></script>
    <script src="/assembly/app.controller.main.js"></script>
    <script src="/assembly/app.run.js"></script>

    <style>
        .width100Percent {
            width: 100%;
        }
    </style>
   
    

    <div ng-app="AsposeAssemblyApp" ng-cloak>
        <div class="container-fluid asposetools assembly" ng-controller="Main">
            <div class="container">
                <div class="row">
                    <div class="col-md-12 pt-5 pb-5">
						<a href="/default" class="btn btn-success btn-lg">Home</a>
                        <h1 runat="server" id="ProductTitle"></h1>
                        <h2 runat="server" id="ProductTitleSub"></h2>

                        <ng-form class="form" ng-show="stage === 1" onsubmit="return false">
                            <div class="uploadfile">
                            <div class="filedropdown">
                                <div class="filedrop">
                                    <label class="dz-message needsclick">
                                        <%= Resources["AssemblyDropTemplate"] %>
                                    </label>
                                    <input type="file"                                           
                                           ngf-select
                                           ng-model="templateFile.file"
                                           ngf-model-invalid="templateError"
                                           ngf-max-size="<%= ViewState["fileMaxSize"] %>"
                                           ngf-pattern="'<%= ViewState["templateValidFileExtensions"] %>'"                                           
                                           accept="<%= ViewState["templateValidFileExtensions"] %>"
                                           required/>
                                    <br/>
                                    <div class="fileupload" ng-show="templateFile.file">
                                        <div class="progress width100Percent" ng-show="templateFile.progress < 100">
                                            <div class="progress-bar progress-bar-striped progress-bar-success progress-bar-animated"
                                                 style="width:{{templateFile.progress}}%"></div>
                                        </div>
                                        <span class="filename" ng-show="templateFile.file.name">
                                            <a ng-click="templateFile={}">
                                                <label class="custom-file-upload">{{templateFile.file.name}}</label>
                                                <i class="fa fa-times">&nbsp;</i>
                                            </a>
                                        </span>
                                    </div>
                                    <br/>
                                    <div runat="server" id="TemplateInvalidExtension" ng-if="templateError.$errorMessages.pattern" class="alert alert-danger" role="alert"></div>
                                    <div runat="server" id="TemplateInvalidSize" ng-if="templateError.$errorMessages.maxSize" class="alert alert-danger" role="alert"></div>
                                </div>

                                <div class="convertbtn">
                                    <a ng-click="uploadTemplateFile()" class="btn btn-success btn-lg"
                                            ng-disabled="!templateFile.file">
                                        <i class="fa fa-upload">&nbsp;</i>Upload Template
                                    </a>
                                </div>
                                <div class="convertbtn">
                                    <a class="btn btn-success btn-lg px-5"
                                            data-toggle="modal" data-target="#help-dialog-template">
                                        <i class="fa fa-info-circle">&nbsp;</i>Help
                                    </a>
                                </div>
                            </div>
                                <div class="clearfix">&nbsp;</div>                                
                                <a ng-click="showStage(2)" class="btn btn-link btn-navigate" style="color:white" ng-if="showed >= 2"
                                    ng-disabled="!templateFile.file || !templateFile.progress || templateFile.progress < 100">
                                    Next&nbsp;<i class="fa fa-arrow-right"></i>
                                </a>
                        </div>
                    </ng-form>

                        <ng-form class="form" ng-show="stage === 2" onsubmit="return false">
                        <div class="uploadfile">
                            <div class="filedropdown">
                                <div class="filedrop">
                                    <label class="dz-message needsclick">
                                        <%= Resources["AssemblyDropDataSource"] %>                                        
                                    </label>
                                    <input type="file"
                                           ngf-select
                                           ng-model="datasourceFile.file"
                                           ngf-model-invalid="datasourceError"
                                           ngf-max-size="<%= ViewState["fileMaxSize"] %>"
                                           ngf-pattern="'<%= ViewState["dataValidFileExtensions"] %>'"                                           
                                           accept="<%= ViewState["dataValidFileExtensions"] %>"                                           
                                           required/>
                                    <br/>

                                    <div class="fileupload" ng-show="datasourceFile.file">
                                        <div class="progress width100Percent" ng-show="datasourceFile.progress < 100">
                                            <div class="progress-bar progress-bar-striped progress-bar-success progress-bar-animated"
                                                 style="width:{{datasourceFile.progress}}%"></div>
                                        </div>
                                        <span class="filename" ng-show="datasourceFile.file.name">
                                            <a ng-click="datasourceFile={};showDelimiter=false;showTableIndex=false">
                                                <label class="custom-file-upload">{{datasourceFile.file.name}}</label>
                                                <i class="fa fa-times">&nbsp;</i>
                                            </a>
                                        </span>
                                    </div>
                                    <br/>
                                    <div runat="server" id="DataInvalidExtension" ng-if="datasourceError.$errorMessages.pattern" class="alert alert-danger" role="alert"></div>
                                    <div runat="server" id="DataInvalidSize" ng-if="datasourceError.$errorMessages.maxSize" class="alert alert-danger" role="alert"></div>
                                </div>
                                <div class="form-inline">
                                    <div class="color-wrapper">                                        
                                        <em>Data Source Name<sup>*</sup></em> <input class="form-control" ng-model="datasourceName" type="text" placeholder="" required/>
                                        <div ng-show="showTableIndex">
                                            <em class="btn">Table Index</em> <input class="form-control" ng-model="datasourceTableIndex" id="datasourceTableIndex" type="number" min="0" style="width:60px"/>                                        
                                        </div>
                                        <div ng-show="showDelimiter">
                                            <em class="btn">Delimiter</em>                                        
                                            <div class="dropdown" style="display:inline-block" id="delimiterdropdown">
                                                <button type="button" class="btn dropdown-toggle" id="delimiter" data-toggle="dropdown" 
                                                    aria-haspopup="true" aria-expanded="false" style="background-color:white;width:60px;border-radius:4px">,</button>
                                                <ul class="dropdown-menu dropdown-menu-left" aria-labelledby="delimiter">                                                
                                                    <li><a ng-click="setDelimiter(',')" class="dropdown-item">,</a></li>
                                                    <li><a ng-click="setDelimiter(';')" class="dropdown-item">;</a></li>
                                                    <li><a ng-click="setDelimiter('\t','Tab')" class="dropdown-item">Tab</a></li>
                                                    <li><a ng-click="setDelimiter(' ','Space')" class="dropdown-item">Space</a></li>
                                                </ul>
                                            </div>                 
                                        </div>
                                        <br/>                                        
                                    </div>
                                </div>
                                <br/>
                                <div class="convertbtn">
                                    <a ng-click="uploadDatasourceFile()" class="btn btn-success btn-lg" 
                                          ng-disabled="!datasourceFile.file || !datasourceName">
                                        <i class="fa fa-upload">&nbsp;</i>Upload Data Source
                                    </a>
                                </div>
                                <div class="convertbtn">
                                    <a class="btn btn-success btn-lg px-5"
                                            data-toggle="modal" data-target="#help-dialog-datasource">
                                        <i class="fa fa-info-circle">&nbsp;</i>Help
                                    </a>
                                </div>
                            </div>                            
                            <div class="clearfix">&nbsp;</div>
                            <a ng-click="showStage(1)" class="btn btn-link btn-navigate" style="color:white">
                                <i class="fa fa-arrow-left">&nbsp;</i>Back
                            </a>
                            <a ng-click="showStage(3)" class="btn btn-link btn-navigate" style="color:white" ng-if="showed >= 3" 
                                ng-disabled="!datasourceFile.file || !datasourceName || !datasourceFile.progress || datasourceFile.progress < 100">
                                Next&nbsp;<i class="fa fa-arrow-right"></i>
                            </a>
                        </div>
                    </ng-form>

                    <ng-form class="form" ng-show="stage === 3" onsubmit="return false">
                        <br/>
                        <div class="uploadfile">         
                            <div class="filedropdown">
                                <div class="filesuccess">         
                                    <label class="dz-message needsclick">
                                        <%= Resources["AssemblyReady"] %>
                                    </label>
                                    <span class="downloadbtn convertbtn">                                    
                                        <a class="btn btn-success btn-lg" ng-click="assembleDocument()" ng-disabled="isLoading">
                                            <i ng-show="!isLoading" class="fa fa-download">&nbsp;</i>
                                            <md-progress-circular md-mode="indeterminate" ng-show="isLoading" md-diameter="18px" class="md-primary">&nbsp;</md-progress-circular>
                                            &nbsp;<%= Resources["AssemblyButton"] %>
                                        </a>                                    
                                    </span>                                    
                                    <br/>                                                            
                                </div>
                            </div>                                                       
                            <br/>
                            <a ng-click="showStage(2)" class="btn btn-link btn-navigate" style="color:white">
                                <i class="fa fa-arrow-left">&nbsp;</i>Back
                            </a>
                            <a ng-click="start()" class="btn btn-link refresh-c btn-navigate" style="color:white">
                                <i class="fa fa-refresh">&nbsp;</i><%=Resources["AssemblyAnotherFile"]%>
                            </a>                            
                        </div>
                    </ng-form>

                        <br />
                        <div id="AssemblyMessage" class="alert" role="alert"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="help-dialog-template" class="modal fade" tabindex="-1" role="dialog" style="z-index: 99999999;">
            <div class="modal-dialog" role="document">
                <div class="modal-content"></div>
            </div>
        </div>

        <div id="help-dialog-datasource" class="modal fade" tabindex="-1" role="dialog" style="z-index: 99999999;">
            <div class="modal-dialog" role="document">
                <div class="modal-content"></div>
            </div>
        </div>
    </div>
    
    <div class="col-md-12 pt-5 app-product-section tl" id="dvAppProductSection" runat="server">
        <div class="container">
            <div class="col-md-3 pull-right">
                <img runat="server" id="ProductImage" />
            </div>
            <div class="col-md-9 pull-left">
                <h3 runat="server" id="AsposeProductTitle"></h3>
                <ul>
                    <li><%=" " + Resources["Supported"] + " " + Resources["Documents"] + ": " + ViewState["validFileExtensions"].ToString().Replace(" or ", ", ")%></li>
                    <li><% = Resources["AssemblyLiFeature1"] %> </li>
                </ul>
            </div>
        </div>
    </div>
        
   

      
    


    <div class="col-md-12 pt-5 app-features-section">
        <div class="container tc pt-5">
            <div class="col-md-4">
                <div class="imgcircle fasteasy">
                    <img src="../../img/fast-easy.png" />
                </div>
                <h4><%= Resources["AssemblyFeature1"] %></h4>
                <p><%= Resources["AssemblyFeature1Description"] %></p>
            </div>
            <div class="col-md-4">
                <div class="imgcircle anywhere">
                    <img src="../../img/anywhere.png" />
                </div>
                <h4><%= Resources["AssemblyFeature2"] %></h4>
                <p><%= Resources["AssemblyFeature2Description"] %></p>
            </div>
            <div class="col-md-4">
                <div class="imgcircle quality">
                    <img src="../../img/quality.png" />
                </div>
                <h4><%= Resources["AssemblyFeature3"] %></h4>
                <p><%= Resources["PoweredBy"] %> <a runat="server" target="_blank" id="PoweredBy"></a><%= Resources["QualityDescMetadata"] %>.</p>
            </div>
        </div>
    </div>
       
	<script type="text/javascript">
        'use strict';
        window.ASPOSE_ASSEMBLY_API = '<%= ViewState["ASPOSE_ASSEMBLY_API"]%>';
        window.ASPOSE_PRODUCTNAME = '<%= ViewState["product"] %>';
        window.ASPOSE_FILEDOWNLOADLINK = '<%= ViewState["ASPOSE_FILEDOWNLOADLINK"] %>';
    </script>
</asp:Content>