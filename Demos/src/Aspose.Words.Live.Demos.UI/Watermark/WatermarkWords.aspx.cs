using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Config;
using System.Web;
using System.Globalization;
using System.Text;
using Aspose.Words.Live.Demos.UI.Helpers;
namespace Aspose.Words.Live.Demos.UI
{
  public partial class WatermarkWords : BasePage
  {
		AsposeWordsWatermark _asposeWordsWatermark = new AsposeWordsWatermark();
    protected void Page_Load(object sender, EventArgs e)
    {
      Product = "words";
      Page.Title = Resources[Product + "WatermarkPageTitle"];
      Page.MetaDescription = Resources[Product + "WatermarkMetaDescription"];
	  AsposeProductTitle.InnerText = PageProductTitle + " " + Resources["WatermarkAPPName"];
      ProductTitle.InnerText = Resources[Product + "WatermarkTitle"];
      ProductTitleSub.InnerText = Resources[Product + "WatermarkTitleSub"];
      ProductImage.Src = "~/img/aspose-" + Product + "-app.png";
      PoweredBy.InnerText = PageProductTitle + ". ";
      PoweredBy.HRef = "https://products.aspose.com/" + Product;
      TextWatermarkButton.Text = Resources["TextWatermarkButton"];
      ProcessTextWatermarkButton.Text = Resources["TextWatermarkButton"];
      ImageWatermarkButton.Text = Resources["ImageWatermarkButton"];
      ProcessImageWatermarkButton.Text = Resources["ImageWatermarkButton"];
      RemoveWatermarkButton.Text = Resources["RemoveWatermarkButton"];
      textWatermark.Attributes.Add("placeholder", Resources["AddWatermarkTextPlaceholder"]);
      
			ViewState["AddanotherWatermark"] = HttpContext.Current.Request.Url.AbsoluteUri;
			ViewState["product"] = Product;
			// Check for auto-generate URLs to set only format as valid extension
			if (Page.RouteData.Values["Format"] != null)
			{
				ViewState["validFileExtensions"] = SetValidation("." + Page.RouteData.Values["Format"].ToString().ToLower(), ValidateFileType);
			}
			else
			{
				ViewState["validFileExtensions"] = SetValidation(Resources[Product + "WatermarkValidationExpression"], ValidateFileType);
			}
      SetValidation(Resources[Product + "WatermarkImageValidationExpression"], ValidateImageType);
      CheckReturnFromViewer(ShowDownloadPage);
			
		}

    protected void UploadFile(Action<FileUploadResponse> action)
    {
      if (IsValid)
        if (CheckFileInputs(UploadFileInput))
          try
          {
            var files = UploadFiles(UploadFileInput);
            if (files != null && files.Count == 1)
            {
              FileNameHidden.Value = files[0].FileName;
              FolderNameHidden.Value = files[0].FolderId;
              action(files[0]);
            }
          }
          catch (Exception ex)
          {
            ShowErrorMessage(WatermarkMessage, "Error: " + ex.Message);
          }
        else
          ShowErrorMessage(WatermarkMessage, Resources["FileSelectMessage"]);
    }

    protected void TextWatermarkButton_Click(object sender, EventArgs e)
    {
      UploadFile((file) =>
      {
        UploadFilePlaceHolder.Visible = false;
        TextPlaceHolder.Visible = true;
        ProductTitleSub.InnerText = Resources[Product + "WatermarkTextTitleSub"];
      });
    }

    protected void ProcessTextWatermarkButton_Click(object sender, EventArgs e)
    {
      
      var response = _asposeWordsWatermark.TextWatermark(FileNameHidden.Value, FolderNameHidden.Value, textWatermark.Value,

		  pickcolor.Value, fontFamily.SelectedValue, double.Parse( fontSize.Text), double.Parse( textAngle.Text));
      SuccessLabel.InnerText = Resources["WatermarkAddedSuccessMessage"];
      PerformResponse(response, TextMessage, ShowDownloadPage);
    }

    protected void ImageWatermarkButton_Click(object sender, EventArgs e)
    {
      UploadFile((file) =>
      {
        UploadFilePlaceHolder.Visible = false;
        ImagePlaceHolder.Visible = true;
        ProductTitleSub.InnerText = Resources[Product + "WatermarkImageTitleSub"];
      });
    }

    protected void ProcessImageWatermarkButton_Click(object sender, EventArgs e)
    {
      if (IsValid)
        if (CheckFileInputs(UploadImageInput))
          try
          {
            var files = UploadFiles(UploadImageInput);
            if (files != null && files.Count == 1)
            {							
              var response = _asposeWordsWatermark.ImageWatermark(FileNameHidden.Value, FolderNameHidden.Value, files[0].FileName, files[0].FolderId, greyScale.Checked, double.Parse(zoom.Text), double.Parse(imageAngle.Text));
              SuccessLabel.InnerText = Resources["WatermarkAddedSuccessMessage"];
              PerformResponse(response, ImageMessage, ShowDownloadPage);
            }
          }
          catch (Exception ex)
          {
            ShowErrorMessage(ImageMessage, "Error: " + ex.Message);
          }
        else
          ShowErrorMessage(ImageMessage, Resources["FileSelectMessage"]);
    }

    protected void RemoveWatermarkButton_Click(object sender, EventArgs e)
    {
      UploadFile((file) =>
      {
        var response = _asposeWordsWatermark.RemoveWatermark(file.FileName, file.FolderId);
        SuccessLabel.InnerText = Resources["WatermarkRemovedSuccessMessage"];
        PerformResponse(response, WatermarkMessage, ShowDownloadPage);
      });
    }

    private void ShowDownloadPage(Response response)
    {
      var url = response.DownloadURL();
      var callbackURL = HttpContext.Current.Request.Url.AbsolutePath;
      var viewerURL = response.ViewerURL(Product, callbackURL);
      DownloadButton.NavigateUrl = url;
     // DownloadUrlInputHidden.Value = HttpUtility.UrlEncode(url);
      ViewerLink.NavigateUrl = viewerURL;
      UploadFilePlaceHolder.Visible = false;
      TextPlaceHolder.Visible = false;
      ImagePlaceHolder.Visible = false;
      DownloadPlaceHolder.Visible = true;
    }

    
  }
}
