<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="WatermarkWords.aspx.cs" Inherits="Aspose.Words.Live.Demos.UI.WatermarkWords" %>

<asp:Content ID="WatermarkContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:UpdatePanel ID="WatermarkPanel" runat="server">
        <ContentTemplate>
            <div class="container-fluid asposetools">
                <div class="container">
                    <div class="row">
                        <div class="col-md-12 pt-5 pb-5">
							<a href="/default" class="btn btn-success btn-lg">Home</a>
                            <h1 runat="server" id="ProductTitle"></h1>
                            <h4 runat="server" id="ProductTitleSub"></h4>
                            <div class="form">
                                <asp:PlaceHolder ID="UploadFilePlaceHolder" runat="server">
                                    <div class="uploadfile">
                                        <div class="filedropdown">
                                            <div class="filedrop">
                                                <asp:HiddenField runat="server" ID="FileNameHidden" ClientIDMode="Static" Value="" />
                                                <asp:HiddenField runat="server" ID="FolderNameHidden" ClientIDMode="Static" Value="" />

                                                <label class="dz-message needsclick"><%=Resources["DropAFile"] %></label>
                                                <input type="file" name="UploadFileInput" id="UploadFileInput" runat="server" class="uploadfileinput" />
                                                <asp:RegularExpressionValidator ID="ValidateFileType" ValidationExpression="([a-zA-Z0-9\s)\s(\s_\\.\-:])+(.doc|.docx|.dot|.dotx|.rtf)$"
                                                    ControlToValidate="UploadFileInput" runat="server" ForeColor="Red"
                                                    Display="Dynamic" />
                                                <div class="fileupload">
                                                    <span class="filename"><a onclick="removefile()">
                                                        <label for="UploadFileInput" class="custom-file-upload"></label>
                                                        <i class="fa fa-times"></i></a></span>
                                                </div>
                                            </div>
                                            <div>
                                                <asp:UpdateProgress ID="UpdateProgressUpload" runat="server" AssociatedUpdatePanelID="WatermarkPanel">
                                                    <ProgressTemplate>
                                                        <div style="padding-bottom: 20px; padding-top: 0px;">
                                                            <img height="59px" width="59px" alt="Please wait..." src="../../img/loader.gif" />
                                                        </div>
                                                    </ProgressTemplate>
                                                </asp:UpdateProgress>
                                            </div>
                                            <p runat="server" id="WatermarkMessage"></p>
                                            <div class="convertbtn" style="padding: 10px;">
                                                <asp:Button runat="server" ID="TextWatermarkButton" class="btn btn-success btn-lg" Text="ADD TEXT WATERMARK" OnClick="TextWatermarkButton_Click"></asp:Button>
                                            </div>
                                            <div class="convertbtn" style="padding: 10px;">
                                                <asp:Button runat="server" ID="ImageWatermarkButton" class="btn btn-success btn-lg" Text="ADD IMAGE WATERMARK" OnClick="ImageWatermarkButton_Click"></asp:Button>
                                            </div>                                            
                                            <div class="convertbtn" style="padding: 10px;">
                                                <asp:Button runat="server" ID="RemoveWatermarkButton" class="btn btn-success btn-lg" Text="REMOVE WATERMARK" OnClick="RemoveWatermarkButton_Click"></asp:Button>
                                            </div>
                                        </div>
                                    </div>
                                </asp:PlaceHolder>
                                <asp:PlaceHolder ID="TextPlaceHolder" runat="server" Visible="false">
                                    <div class="watermark" style="margin-bottom: 0px; margin-top: 30px;">
                                        <textarea id="textWatermark" runat="server" class="form-control" aria-describedby="basic-addon2"></textarea>
                                        <br />
                                        <asp:RequiredFieldValidator ID="rfvWatermark" EnableClientScript="true" runat="server"
                                            ControlToValidate="textWatermark" Display="Dynamic"
                                            ValidationGroup="settingsTextStamp"></asp:RequiredFieldValidator>
                                    </div>
                                    <div class="colorpicker">
                                        <div class="form-inline">
                                            <div class="color-wrapper">
                                                <input type="text" name="custom_color" placeholder="#99FF66" value="#99FF66" clientidmode="Static" id="pickcolor" class="call-picker" runat="server" />
                                                <div class="color-holder call-picker"></div>
                                                <div class="color-picker" id="color-picker" style="display: none"></div>
                                                &nbsp;
					                                <asp:DropDownList CssClass="form-control" ID="fontFamily" runat="server">
                                                        <asp:ListItem Selected="True" Value="Arial">Arial</asp:ListItem>
                                                        <asp:ListItem Value="Times New Roman">Times New Roman</asp:ListItem>
                                                        <asp:ListItem Value="Courier">Courier</asp:ListItem>
                                                        <asp:ListItem Value="Verdana">Verdana</asp:ListItem>
                                                        <asp:ListItem Value="Helvetica">Helvetica</asp:ListItem>
                                                        <asp:ListItem Value="Georgia">Georgia</asp:ListItem>
                                                        <asp:ListItem Value="Comic Sans MS">Comic Sans MS</asp:ListItem>
                                                        <asp:ListItem Value="Trebuchet MS">Trebuchet MS</asp:ListItem>
                                                        <asp:ListItem Value="Calibri">Calibri</asp:ListItem>
                                                    </asp:DropDownList>
                                                &nbsp;
                                                    <asp:TextBox runat="server" ID="fontSize" class="form-control" TextMode="Number" min="8" Style="width: 60px">72</asp:TextBox>
                                            </div>
                                            <div class="form-inline">
                                                <div class="color-wrapper">
                                                    <p style="display: inline-block"><%=Resources["RotateAngleLabel"] %>&nbsp;(-360&deg;&nbsp;..&nbsp;360&deg;)</p>
                                                    &nbsp;
				                                    <asp:TextBox runat="server" ID="textAngle" CssClass="form-control" TextMode="Number" min="-360" max="360" step="45" Style="display: inline-block">-45</asp:TextBox>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    <p runat="server" id="TextMessage"></p>
                                    <div class="convertbtn">
                                        <asp:Button runat="server" ID="ProcessTextWatermarkButton" class="btn btn-success btn-lg" Text="ADD TEXT WATERMARK" OnClick="ProcessTextWatermarkButton_Click"></asp:Button>
                                    </div>
                                </asp:PlaceHolder>

                                <asp:PlaceHolder ID="ImagePlaceHolder" runat="server" Visible="false">
                                    <div class="filedropdown">
                                        <div class="filedrop">
                                            <label class="dz-message needsclick"><%=Resources["DropAFile"] %></label>
                                            <input type="file" name="UploadImageInput" id="UploadImageInput" runat="server" class="uploadfileinput" />
                                            <asp:RegularExpressionValidator ID="ValidateImageType" ValidationExpression="([a-zA-Z0-9\s)\s(\s_\\.\-:])+(.jpg|.bmp|.png)$"
                                                ControlToValidate="UploadImageInput" runat="server" ForeColor="Red"
                                                Display="Dynamic" />
                                            <div class="fileupload">
                                                <span class="filename"><a onclick="removefile()">
                                                    <label for="UploadImageInput" class="custom-file-upload"></label>
                                                    <i class="fa fa-times"></i></a></span>
                                            </div>
                                        </div>
                                        
                                    </div>
                                    
                                    <div class="watermark" style="margin-bottom: 0px;" runat="server">
                                        <div class="form-inline">
                                            <div class="color-wrapper">                                            
                                                <p style="display: inline-block;"><%=Resources["GrayscaledLabel"] %></p>
                                                &nbsp;      
                                                <asp:CheckBox id="greyScale" runat="server"/>                                                              
                                                </div>
                                        <div class="color-wrapper">                                            
                                            <p style="display: inline-block;"><%=Resources["ZoomFactorLabel"] %>&nbsp; (&#37;)</p>
                                            &nbsp;                                            
                                            <asp:TextBox runat="server" ID="zoom" CssClass="form-control" TextMode="Number" min="0" max="400">100</asp:TextBox>                                            
                                        </div>                                        
                                            <div class="color-wrapper">
                                                <p style="display: inline-block;"><%=Resources["RotateAngleLabel"] %>&nbsp; (-360&deg;&nbsp;..&nbsp;360&deg;)</p>
                                                &nbsp;
                                                <asp:TextBox runat="server" ID="imageAngle" CssClass="form-control" TextMode="Number" min="-360" max="360">0</asp:TextBox>
                                            </div>                                        
                                        </div>
                                    </div>                                    
                                    
                                    <div>                                        
                                        <asp:UpdateProgress ID="UpdateProgressImage" runat="server" AssociatedUpdatePanelID="WatermarkPanel">
                                            <ProgressTemplate>
                                                <div style="padding-bottom: 20px; padding-top: 0px;">
                                                    <img height="59px" width="59px" alt="Please wait..." src="../../img/loader.gif" />
                                                </div>
                                            </ProgressTemplate>
                                        </asp:UpdateProgress>
                                    </div>
                                    <p runat="server" id="ImageMessage"></p>
                                    <div class="convertbtn" style="margin-top: 10px">
                                        <asp:Button runat="server" ID="ProcessImageWatermarkButton" class="btn btn-success btn-lg" Text="ADD IMAGE WATERMARK" OnClick="ProcessImageWatermarkButton_Click"></asp:Button>
                                    </div>                                    
                                </asp:PlaceHolder>

                                <asp:PlaceHolder ID="DownloadPlaceHolder" runat="server" Visible="false">
                                    <div class="filesendemail">
                                        <div class="filesuccess">
                                            <label class="dz-message needsclick" id="SuccessLabel" runat="server"></label>
                                            <span class="downloadbtn convertbtn">
                                                <asp:HyperLink NavigateUrl="#" ID="DownloadButton" Target="_blank" runat="server" CssClass="btn btn-success btn-lg"><%= Resources["DownLoadNow"] %> <i class="fa fa-download"></i></asp:HyperLink>
                                            </span>
                                            <div class="clearfix">&nbsp;</div>
                                            <span class="viewerbtn">
                                                <asp:HyperLink NavigateUrl="#" ID="ViewerLink" Target="_self" runat="server" CssClass="btn btn-success btn-lg"><%=Resources["WatermarkViewer"]%> <i class="fa fa-eye"></i></asp:HyperLink>
                                            </span>
                                            <div class="clearfix">&nbsp;</div>
                                            <a href="<%= ViewState["AddanotherWatermark"]%>" class="btn btn-link refresh-c"><%=Resources["WatermarkAnotherFile"]%> <i class="fa-refresh fa "></i></a>
                                        </div>
                                        
                                        <br />
                                        
                                        
                                        <asp:UpdateProgress ID="UpdateProgressDownload" runat="server" AssociatedUpdatePanelID="WatermarkPanel">
                                            <ProgressTemplate>
                                                <div>
                                                    <img height="59px" width="59px" alt="Please wait..." src="../../img/loader.gif" />
                                                </div>
                                            </ProgressTemplate>
                                        </asp:UpdateProgress>
                                        <p runat="server" id="DownloadMessage"></p>
                                    </div>
                                </asp:PlaceHolder>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </ContentTemplate>
        <Triggers>
            <asp:PostBackTrigger ControlID="TextWatermarkButton" />
            <asp:PostBackTrigger ControlID="ProcessTextWatermarkButton" />
            <asp:PostBackTrigger ControlID="ImageWatermarkButton" />
            <asp:PostBackTrigger ControlID="ProcessImageWatermarkButton" />
            <asp:PostBackTrigger ControlID="RemoveWatermarkButton" />
        </Triggers>
    </asp:UpdatePanel>
    <div class="col-md-12 pt-5 app-product-section tl" id="dvAppProductSection" runat="server">
        <div class="container">
            <div class="col-md-3 pull-right">
                <img runat="server" id="ProductImage" />
            </div>
            <div class="col-md-9 pull-left">
                <h3 runat="server" id="AsposeProductTitle"></h3>
                <ul>
                    <li><%=" " + Resources["Supported"] + " " + Resources["Documents"] + ": " + ViewState["validFileExtensions"].ToString().Replace(" or ", ", ")%></li>
                    <li><%= Resources["WordsWatermarkLiFeature1"] %></li>
                </ul>
            </div>
        </div>
    </div>
	

<!-- HowTo Section -->
        <div class="col-md-12 tl bg-darkgray howtolist"  id="dvHowToSection" visible="false" runat="server"><div class="container tl dflex">
        
            <div class="col-md-4 howtosectiongfx"><img src="https://products.aspose.app/img/howto.png"  ></div><div class="howtosection col-md-8"><div><h4><i class="fa fa-question-circle "></i> <b><%= string.Format(Resources["HowtoWatermarkTitle"], ViewState["Extension1"], AsposeProductTitle.InnerText) %></b></h4><ul>
				<li><%= string.Format(Resources["HowtoWatermarkFeature1"], ViewState["Extension1"], ViewState["Extension1"]) %></li>
				<li><%= string.Format(Resources["HowtoWatermarkFeature2"], ViewState["Extension1"].ToString()) %></li>
				<li><%= string.Format(Resources["HowtoWatermarkFeature3"], ViewState["Extension1"].ToString()) %></li>
				<li><%= string.Format(Resources["HowtoWatermarkFeature4"], ViewState["Extension1"].ToString()) %></li>
				<li><%= Resources["HowtoWatermarkFeature5"]%></li> </ul></div></div></div></div>
    <div class="col-md-12 pt-5 app-features-section">
        <div class="container tc pt-5">
            <div class="col-md-4">
                <div class="imgcircle fasteasy">
                    <img src="../../img/fast-easy.png" />
                </div>
                <h4><%= Resources["WordsWatermarkFeature1"] %></h4>
                <p><%= Resources["WordsWatermarkFeature1Description"] %></p>
            </div>
            <div class="col-md-4">
                <div class="imgcircle anywhere">
                    <img src="../../img/anywhere.png" />
                </div>
                <h4><%= Resources["WordsWatermarkFeature2"] %></h4>
                <p><%= Resources["WordsWatermarkFeature2Description"] %>.</p>
            </div>
            <div class="col-md-4">
                <div class="imgcircle quality">
                    <img src="../../img/quality.png" />
                </div>
                <h4><%= Resources["WordsWatermarkFeature3"] %></h4>
                <p><%= Resources["PoweredBy"] %> <a runat="server" target="_blank" id="PoweredBy"></a><%= Resources["QualityDescMetadata"] %>.</p>
            </div>
        </div>
    </div>
	
    <script type="text/javascript">
        window.onsubmit = function () {
            if (Page_IsValid) {
                var updateProgressUpload = $find("<%= UpdateProgressUpload.ClientID %>");
                if (updateProgressUpload) {
                    window.setTimeout(function () {
                        updateProgressUpload.set_visible(true);
                        document.getElementById('<%= WatermarkMessage.ClientID %>').style.display = 'none';
                    }, 100);
                }
                var updateProgressDownload = $find("<%= UpdateProgressDownload.ClientID %>");
                if (updateProgressDownload) {
                    window.setTimeout(function () {
                        updateProgressDownload.set_visible(true);
                        document.getElementById('<%= DownloadMessage.ClientID %>').style.display = 'none';
                    }, 100);
                }
            }
        }
    </script>
    <script>
        $('.fileupload').hide();
        $('.uploadfileinput').change(function () {
            $('.fileupload').hide();
            var file = $('.uploadfileinput')[0].files[0].name;
            $('.filename label').text(file);
            $('.fileupload').show();
            var message = $('#<%= WatermarkMessage.ClientID %>');
            if (message.length)
                message[0].style.display = 'none';
            var imageMessage = $('#<%= ImageMessage.ClientID %>');
            if (imageMessage.length)
                imageMessage[0].style.display = 'none';
        });
        function removefile() {
            $('.fileupload').hide();
            $('.uploadfileinput').show();
        }
    </script>

     <link rel="stylesheet" href="https://products.aspose.app/css/colorpicker.css" type="text/css" />
    <script src="https://products.aspose.app/js/colorpicker.js"></script>
</asp:Content>
