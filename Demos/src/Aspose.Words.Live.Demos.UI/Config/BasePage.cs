using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Aspose.Words.Live.Demos.UI.Services;
using Aspose.Words.Live.Demos.UI.Models;

namespace Aspose.Words.Live.Demos.UI.Config
{
	public class BasePage : BaseRootPage
	{
		private string _product;
		/// <summary>
		/// Product name (e.g. words, cells)
		/// </summary>
		public string Product
		{
			get
			{
				if (string.IsNullOrEmpty(_product))
				{
					if (Page.RouteData.Values["Product"] != null)
					{
						_product = Page.RouteData.Values["Product"].ToString().ToLower();
					}
				}		
				return _product;
			}
			set => _product = value;
		}
		public string _pageProductTitle;
		/// <summary>
		/// Product title (e.g. Aspose.Words)
		/// </summary>
		public string PageProductTitle
		{
			get
			{
				if (_pageProductTitle == null)
					_pageProductTitle = Resources["Aspose" + TitleCase(Product)];
				return _pageProductTitle;
			}
		}

		private int _appURLID = 0;
		public int AppURLID
		{
			get
			{
				return _appURLID;
			}
		}

		protected override void OnLoad(EventArgs e)
		{
			if (Resources != null)
			{
				Page.Title = Resources["ApplicationTitle"];
			}

			base.OnLoad(e);
		}

		/// <summary>
		/// Set validation for RegularExpressionValidators, InputFile and ViewStates using Resources[Product + "ValidationExpression"]
		/// </summary>
		private void SetAccept(string validationExpression, params HtmlInputFile[] inpufiles)
		{
			var accept = validationExpression.ToLower().Replace("|", ",");
			foreach (var inpufile in inpufiles)
				inpufile.Attributes.Add("accept", accept.ToLower().Replace("|", ","));
		}

		/// <summary>
		/// Set validation for RegularExpressionValidators and ViewStates using Resources[Product + "ValidationExpression"].
		/// If the 'ControlToValidate' option is HtmlInputFile, it sets accept an attribute to that control
		/// </summary>
		/// <param name="validators"></param>
		protected void SetValidation(params RegularExpressionValidator[] validators)
		{

			var validationExpression = "";
			var validFileExtensions = "";
			// Check for format like .Doc
			if (Page.RouteData.Values["Format"] != null)
			{
				validFileExtensions = Page.RouteData.Values["Format"].ToString().ToUpper();
				validationExpression = "." + validFileExtensions.ToLower();
			}
			else
			{
				validationExpression = Resources[Product + "ValidationExpression"];
				validFileExtensions = GetValidFileExtensions(validationExpression);
			}

			SetValidation(validationExpression, validators);

			ViewState["product"] = Product;
			ViewState["validFileExtensions"] = validFileExtensions;
		}

		protected void SetValidationExpression(string validationExpression, params RegularExpressionValidator[] validators)
		{
			var validFileExtensions = GetValidFileExtensions(validationExpression);
			SetValidation(validationExpression, validators);

			ViewState["product"] = Product;
			ViewState["validFileExtensions"] = validFileExtensions;
		}

		/// <summary>
		/// Set validation for RegularExpressionValidators, InputFile and ViewStates using validationExpression
		/// </summary>
		protected string SetValidation(string validationExpression, params RegularExpressionValidator[] validators)
		{

			// Check for format if format is available then valid expression will be only format for auto generated URLs
			if (Page.RouteData.Values["Format"] != null)
			{

				validationExpression = "." + Page.RouteData.Values["Format"].ToString().ToLower();
			}

			var validFileExtensions = GetValidFileExtensions(validationExpression);

			foreach (var v in validators)
			{
				v.ValidationExpression = @"(.*?)(" + validationExpression.ToLower() + "|" + validationExpression.ToUpper() + ")$";
				v.ErrorMessage = Resources["InvalidFileExtension"] + " " + validFileExtensions;
				if (!string.IsNullOrEmpty(v.ControlToValidate))
				{
					var control = v.FindControl(v.ControlToValidate);
					if (control is HtmlInputFile inpufile)
						SetAccept(validationExpression, inpufile);
				}
			}
			return validFileExtensions;
		}

		/// <summary>
		/// Get the text of valid file extensions
		/// e.g. DOC, DOCX, DOT, DOTX, RTF or ODT
		/// </summary>
		/// <param name="validationExpression"></param>
		/// <returns></returns>
		protected string GetValidFileExtensions(string validationExpression)
		{
			var validFileExtensions = validationExpression.Replace(".", "").Replace("|", ", ").ToUpper();

			var index = validFileExtensions.LastIndexOf(",");
			if (index != -1)
			{
				var substr = validFileExtensions.Substring(index);
				var str = substr.Replace(",", " or");
				validFileExtensions = validFileExtensions.Replace(substr, str);
			}

			return validFileExtensions;
		}

		

		/// <summary>
		/// Check for null and ContentLength of the PostedFile
		/// </summary>
		/// <param name="fileInputs"></param>
		/// <returns></returns>
		protected bool CheckFileInputs(params HtmlInputFile[] fileInputs)
		{
			return fileInputs.All(x => x != null && x.PostedFile.ContentLength > 0);
		}

		/// <summary>
		/// Save uploaded file to the directory
		/// </summary>
		/// <returns>SaveLocation with filename</returns>
		private FileUploadResponse SaveUploadedFile(string directory, HtmlInputFile fileInput, string folder)
		{
			var fn = Path.GetFileName(fileInput.PostedFile.FileName); // Edge browser sents a full path for a filename
			var saveLocation = Path.Combine(directory, fn);
			fileInput.PostedFile.SaveAs(saveLocation);

			return new FileUploadResponse
			{
				FileName = fn,
				FileLength = fileInput.PostedFile.ContentLength,
				FolderId = folder,
				LocalFilePath = saveLocation
			};
		}
		/// <summary>
		/// Check response for null and StatusCode. Call action if everything is fine or show error message if not
		/// </summary>
		/// <param name="response"></param>
		/// <param name="controlErrorMessage"></param>
		/// <param name="action"></param>
		protected void PerformResponse(Response response, HtmlGenericControl controlErrorMessage, Action<Response> action)
		{
			if (response == null)
				throw new Exception(Resources["ResponseTime"]);

			if (response.StatusCode == 200 && response.FileProcessingErrorCode == 0)
				action(response);
			else
				ShowErrorMessage(controlErrorMessage, response);
		}

		/// <summary>
		/// Check FileProcessingErrorCode of the response and show the error message
		/// </summary>
		/// <param name="control"></param>
		/// <param name="response"></param>
		protected void ShowErrorMessage(HtmlGenericControl control, Response response)
		{
			string txt;
			switch (response.FileProcessingErrorCode)
			{
				case FileProcessingErrorCode.NoImagesFound:
					txt = Resources["NoImagesFoundMessage"];
					break;
				case FileProcessingErrorCode.NoSearchResults:
					txt = Resources["NoSearchResultsMessage"];
					break;
				case FileProcessingErrorCode.WrongRegExp:
					txt = Resources["WrongRegExpMessage"];
					break;
				default:
					txt = response.Status;
					break;
			}
			ShowErrorMessage(control, txt);
		}
		protected Collection<FileUploadResponse> UploadFiles(params HtmlInputFile[] fileInputs)
		{
			//Collection<FileUploadResponse> uploadResult = null;
			string _folderName = Guid.NewGuid().ToString();
			var directory = Path.Combine(Configuration.WorkingDirectory,_folderName);
			Directory.CreateDirectory(directory);
			var uploadResult =  fileInputs.Select(x => SaveUploadedFile(directory, x, _folderName)).ToArray();

			var list = new Collection<FileUploadResponse>(uploadResult);
			return list;
		}
		protected void ShowErrorMessage(HtmlGenericControl control, string message)
		{
			if(message.ToLower().Contains("password"))
			{		
				if ("pdf words cells slides note".Contains(Product.ToLower()) && !AppRelativeVirtualPath.ToLower().Contains("unlock"))
				{
					string asposeUrl = "/" + Product + "/unlock";
					message = "Your file seems to be encrypted. Please use our <a style=\"color:yellow\" href=\"" + asposeUrl + "\">" + PageProductTitle + " Unlock</a> app to remove the password.";
				}
			}

			control.Visible = true;
			control.InnerHtml = message;
			control.Attributes.Add("class", "alert alert-danger");
		}

		protected void ShowSuccessMessage(HtmlGenericControl control, string message)
		{
			control.Visible = true;
			control.InnerHtml = message;
			control.Attributes.Add("class", "alert alert-success");
		}

		protected void CheckReturnFromViewer(Action<Response> action)
		{
			if (Request.QueryString["folderName"] != null && Request.QueryString["fileName"] != null)
			{
				var response = new Response()
				{
					FileName = Request.QueryString["fileName"],
					FolderName = Request.QueryString["folderName"]
				};
				action(response);
			}
		}

		protected string TitleCase(string value)
		{
			return new CultureInfo("en-US", false).TextInfo.ToTitleCase(value);
		}
	}
}
