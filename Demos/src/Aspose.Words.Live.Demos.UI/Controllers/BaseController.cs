using Aspose.Words.Live.Demos.UI.Config;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Services;
using Aspose.Words.Live.Demos.UI.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net.Http;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public abstract class BaseController : Controller
	{
		
		protected Document[] UploadDocuments(HttpRequestBase Request, string sourceFolder)
		{
			try
			{								
				var pathProcessor = new PathProcessor(sourceFolder);
				List<Document> documents = new List<Document>();
				//foreach (string fileName in Request.Files)
				for (int i = 0; i < Request.Files.Count; i++)
				{
					HttpPostedFileBase postedFile = Request.Files[i];

					if (postedFile != null)
					{
						// Check if File is available.
						if (postedFile != null && postedFile.ContentLength > 0)
						{
							string _savepath = pathProcessor.SourceFolder + "\\" + System.IO.Path.GetFileName(postedFile.FileName);
							postedFile.SaveAs(_savepath);
							documents.Add(new Document(_savepath));
						}
					}
				}
				return documents.ToArray();
			}
			catch (IncorrectPasswordException)
			{
				return null;
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Document[0];
			}
		}

		protected InputFiles UploadFiles(HttpRequestBase Request)
		{
			try
			{
				string _folderName = Guid.NewGuid().ToString();
				var pathProcessor = new PathProcessor(_folderName);
				InputFiles _inputFiles = new InputFiles();
				//foreach (string fileName in Request.Files)
				for (int i = 0; i < Request.Files.Count; i++)

				{
					HttpPostedFileBase postedFile = Request.Files[i];

					if (postedFile != null)
					{
						// Check if File is available.
						if (postedFile != null && postedFile.ContentLength > 0)
						{
							string fn = System.IO.Path.GetFileName(postedFile.FileName);
							string _savepath = pathProcessor.SourceFolder + "\\" + fn;
							postedFile.SaveAs(_savepath);
							_inputFiles.Add(new InputFile(fn, _folderName, _savepath));

						}
					}
				}
				return _inputFiles;
			}
			catch (IncorrectPasswordException)
			{
				return null;
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return null;
			}
		}
		/// <summary>
		/// Response when uploaded files exceed the limits
		/// </summary>
		protected Response BadDocumentResponse = new Response()
		{
			Status = "Some of your documents are corrupted",
			StatusCode = 500
		};


		public abstract string Product { get; }

		protected override void OnActionExecuted(ActionExecutedContext ctx)
		{
			base.OnActionExecuted(ctx);
			ViewBag.Title = ViewBag.Title ?? Resources["ApplicationTitle"];
			ViewBag.MetaDescription = ViewBag.MetaDescription ?? "Save time and software maintenance costs by running single instance of software, but serving multiple tenants/websites. Customization available for Joomla, Wordpress, Discourse, Confluence and other popular applications.";
		}

		private AsposeAppContext _atcContext;


		/// <summary>
		/// Main context object to access all the dcContent specific context info
		/// </summary>
		public AsposeAppContext AsposeToolsContext
		{
			get
			{
				if (_atcContext == null) _atcContext = new AsposeAppContext(HttpContext.ApplicationInstance.Context);
				return _atcContext;
			}
		}

		private Dictionary<string, string> _resources;

		/// <summary>
		/// key/value pair containing all the error messages defined in resources.xml file
		/// </summary>
		public Dictionary<string, string> Resources
		{
			get
			{
				if (_resources == null) _resources = AsposeToolsContext.Resources;
				return _resources;
			}
		}

		protected bool IsValidEmail(string email)
		{
			try
			{
				var addr = new System.Net.Mail.MailAddress(email);
				return addr.Address == email;
			}
			catch
			{
				return false;
			}
		}


	}
}
