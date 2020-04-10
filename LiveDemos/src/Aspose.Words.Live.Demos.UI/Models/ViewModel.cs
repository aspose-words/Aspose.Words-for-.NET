using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using Aspose.Words.Live.Demos.UI.Config;
using Aspose.Words.Live.Demos.UI.Controllers;

namespace Aspose.Words.Live.Demos.UI.Models
{
	public class ViewModel
	{
		public int MaximumUploadFiles { get; set; }

		/// <summary>
		/// Name of the product (e.g., words)
		/// </summary>
		public string Product { get; set; }

		public BaseController Controller;

		/// <summary>
		/// Product + AppName, e.g. wordsMerger
		/// </summary>
		public string ProductAppName { get; set; }
		private AsposeAppContext _atcContext;
		public AsposeAppContext AsposeAppContext
		{
			get
			{
				if (_atcContext == null) _atcContext = new AsposeAppContext(HttpContext.Current);
				return _atcContext;
			}
		}
		private Dictionary<string, string> _resources;
		public Dictionary<string, string> Resources
		{
			get
			{
				if (_resources == null) _resources = AsposeAppContext.Resources;
				return _resources;
			}
			set
			{
				_resources = value;
			}
		}

		public string UIBasePath => Configuration.AsposeAppLiveDemosPath;

		public string PageProductTitle => Resources["Aspose" + TitleCase(Product)];

		/// <summary>
		/// The name of the app (e.g., Conversion, Merger)
		/// </summary>
		public string AppName { get; set; }

		/// <summary>
		/// The full address of the application without query string (e.g., https://products.aspose.app/words/conversion)
		/// </summary>
		public string AppURL { get; set; }

		/// <summary>
		/// File extension without dot received by "fileformat" value in RouteData (e.g. docx)
		/// </summary>
		public string Extension { get; set; }

		/// <summary>
		/// File extension without dot received by "fileformat" value in RouteData (e.g. docx)
		/// </summary>
		public string Extension2 { get; set; }

		/// <summary>
		/// Redirect to main app, if there is no ExtensionInfoModel for auto generated models
		/// </summary>
		public bool RedirectToMainApp { get; set; }

		/// <summary>
		/// Name of the partial View of controls (e.g. UnlockControls)
		/// </summary>
		public string ControlsView { get; set; }

		/// <summary>
		/// Is canonical page opened (/all)
		/// </summary>
		public bool IsCanonical;

		

		public string AnotherFileText { get; set; }
		public string UploadButtonText { get; set; }
		public string ViewerButtonText { get; set; }
		public bool ShowViewerButton { get; set; }
		public string SuccessMessage { get; set; }

		/// <summary>
		/// List of app features for ul-list. E.g. Resources[app + "LiFeature1"]
		/// </summary>
		public List<string> AppFeatures { get; set; }

		public string Title { get; set; }
		public string TitleSub { get; set; }


		public string PageTitle
		{
			get => Controller.ViewBag.PageTitle;
			set => Controller.ViewBag.PageTitle = value;
		}

		public string MetaDescription
		{
			get => Controller.ViewBag.MetaDescription;
			set => Controller.ViewBag.MetaDescription = value;
		}

		public string MetaKeywords
		{
			get => Controller.ViewBag.MetaKeywords;
			set => Controller.ViewBag.MetaKeywords = value;
		}

		/// <summary>
		/// If the application doesn't need to upload several files (e.g. Viewer, Editor)
		/// </summary>
		public bool UploadAndRedirect { get; set; }

		protected string TitleCase(string value) => new System.Globalization.CultureInfo("en-US", false).TextInfo.ToTitleCase(value);

		/// <summary>
		/// e.g., .doc|.docx|.dot|.dotx|.rtf|.odt|.ott|.txt|.html|.xhtml|.mhtml
		/// </summary>
		public string ExtensionsString { get; set; }

		#region SaveAs
		private bool _saveAsComponent;
		public bool SaveAsComponent
		{
			get => _saveAsComponent;
			set
			{
				_saveAsComponent = value;
				Controller.ViewBag.SaveAsComponent = value;
				if (_saveAsComponent)
				{
					var sokey1 = $"{Product}{AppName}SaveAsOptions";
					var sokey2 = $"{Product}SaveAsOptions";
					
					if (Resources.ContainsKey(sokey1))
						SaveAsOptions = Resources[sokey1].Split(',');
					else if (Resources.ContainsKey(sokey2))
					{
						if (AppName == "Conversion" && Product == "words")
						{
							var lst = Resources[sokey2].Split(',').ToList();
							try
							{

								var index = lst.FindIndex(x => x == "DOCX");
								lst.RemoveAt(index);
								var index2 = lst.FindIndex(x => x == "DOC");
								lst.Insert(index2, "DOCX");

							}
							catch
							{
								//
							}
							finally
							{
								SaveAsOptions = lst.ToArray();
							}
						}
						else if (AppName == "Conversion" && Product == "pdf")
						{
							var lst = Resources[sokey2].Split(',').ToList().Select(x => x.ToUpper().Trim()).ToList();
							try
							{

								var index = lst.FindIndex(x => x == "DOCX");
								lst.RemoveAt(index);
								var index2 = lst.FindIndex(x => x == "DOC");
								lst.Insert(index2, "DOCX");

							}
							catch
							{
								//
							}
							finally
							{
								SaveAsOptions = lst.ToArray();
							}
						}
						else if (AppName == "Conversion" && Product == "page")
						{
							var lst = Resources[sokey2].Split(',').ToList().Select(x => x.ToUpper().Trim()).ToList();
							SaveAsOptions = lst.ToArray();

						}
						else
							SaveAsOptions = Resources[sokey2].Split(',');
					}

					var lifeaturekey = Product + "SaveAsLiFeature";
					if (AppFeatures != null && Resources.ContainsKey(lifeaturekey))
						AppFeatures.Add(Resources[lifeaturekey]);
				}
			}
		}

		public string SaveAsOptionsList
		{
			get
			{
				string list = "";
				if (SaveAsOptions != null)
				{
					foreach (var extensin in SaveAsOptions)
					{
						if (list == "")
						{
							list = extensin.ToUpper();
						}
						else
						{
							list = list + ", " + extensin.ToUpper();
						}
					}
				}
				return list;

			}
		}
		/// <summary>
		/// FileFormats in UpperCase
		/// </summary>
		public string[] SaveAsOptions { get; set; }

		/// <summary>
		/// Original file format SaveAs option for multiple files uploading
		/// </summary>
		public bool SaveAsOriginal { get; set; }
		#endregion

		/// <summary>
		/// The possibility of changing the order of uploaded files. It is actual for Merger App.
		/// </summary>
		public bool UseSorting { get; set; }

		public string DropOrUploadFileLabel { get; set; }
		

		#region ViewSections
		public bool ShowExtensionInfo => ExtensionInfoModel != null;
		public ExtensionInfoModel ExtensionInfoModel { get; set; }

		public bool HowTo => HowToModel != null;
		public HowToModel HowToModel { get; set; }

		#endregion

		public string JSOptions => new JSOptions(this).ToString();

		public ViewModel(BaseController controller, string app)
		{
			Controller = controller;
			Resources = controller.Resources;
			AppName = Resources.ContainsKey($"{app}APPName") ? Resources[$"{app}APPName"] : app;
			Product = controller.Product;
			var url = controller.Request.Url.AbsoluteUri;
			AppURL = url.Substring(0, (url.IndexOf("?") > 0 ? url.IndexOf("?") : url.Length));
			ProductAppName = Product + app;

			UploadButtonText = GetFromResources(ProductAppName + "Button", app + "Button");
			ViewerButtonText = GetFromResources(app + "Viewer", "ViewDocument");
			SuccessMessage = GetFromResources(app + "SuccessMessage");
			AnotherFileText = GetFromResources(app + "AnotherFile");


			IsCanonical = true;

			HowToModel = new HowToModel(this);

			SetTitles();
			SetAppFeatures(app);
			ShowViewerButton = true;
			SaveAsOriginal = true;
			SaveAsComponent = false;
			SetExtensionsString();
		}

		private void SetTitles()
		{
			PageTitle = Resources[Product + AppName + "PageTitle"];
			MetaDescription = Resources[Product + AppName + "MetaDescription"];
			MetaKeywords = "";
			Title = Resources[Product + AppName + "Title"];
			TitleSub = Resources[Product + AppName + "SubTitle"];
			Controller.ViewBag.CanonicalTag = null;
		}

		private void SetAppFeatures(string app)
		{

			AppFeatures = new List<string>();

			var i = 1;
			while (Resources.ContainsKey($"{ProductAppName}LiFeature{i}"))
				AppFeatures.Add(Resources[$"{ProductAppName}LiFeature{i++}"]);

			// Stop other developers to add unnecessary features.
			if (AppFeatures.Count == 0)
			{
				i = 1;
				while (Resources.ContainsKey($"{app}LiFeature{i}"))
				{
					if (!Resources[$"{app}LiFeature{i}"].Contains("Instantly download") || AppFeatures.Count == 0)
						AppFeatures.Add(Resources[$"{app}LiFeature{i}"]);
					i++;
				}
			}

		}

		private string GetFromResources(string key, string defaultKey = null)
		{
			if (Resources.ContainsKey(key))
				return Resources[key];
			if (!string.IsNullOrEmpty(defaultKey) && Resources.ContainsKey(defaultKey))
				return Resources[defaultKey];
			return "";
		}

		private void SetExtensionsString()
		{
			if (!ShowExtensionInfo)
			{
				var key1 = $"{Product}{AppName}ValidationExpression";
				var key2 = $"{Product}ValidationExpression";
				ExtensionsString = Resources.ContainsKey(key1) ? Resources[key1] : Resources[key2];
				if ("pdf".Equals(Product))
				{
					switch (Extension)
					{
						case "mhtml":
						case "mht":
							ExtensionsString = ".mht|.mhtml";
							break;
						default:
							if (!IsCanonical)
								ExtensionsString = $".{Extension}";
							else
								ExtensionsString = $".pdf";
							break;
					}
				}
			}
			else
			{
				switch (Extension)
				{
					case "doc":
					case "docx":
						ExtensionsString = ".docx|.doc";
						break;
					case "html":
					case "htm":
					case "mhtml":
					case "mht":
						ExtensionsString = ".htm|.html|.mht|.mhtml";
						break;
					default:
						ExtensionsString = $".{Extension}";
						break;
				}

				if (AppName == "Comparison" && !string.IsNullOrEmpty(Extension2))
					ExtensionsString += $"|.{Extension2}";
			}
		}
	}
}
