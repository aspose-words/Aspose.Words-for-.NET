using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class RedactionController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Redaction(string outputType, string searchQuery, string replaceText, bool caseSensitive, bool text, bool comments, bool metadata)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadDocuments(Request, _sourceFolder);

				AsposeWordsRedaction asposeWordsRedaction = new AsposeWordsRedaction();
				response = asposeWordsRedaction.Redact(docs, _sourceFolder, outputType, searchQuery, replaceText, caseSensitive, text, comments, metadata);

			}

			return response;				
		}
		public ActionResult Redaction()
		{
			var model = new ViewModel(this, "Redaction")
			{
				ControlsView = "RedactionControls",
				SaveAsComponent = true,
				MaximumUploadFiles = 10,
				DropOrUploadFileLabel = Resources["DropOrUploadFiles"]
			};
			if (model.RedirectToMainApp)
				return Redirect("/words/" + model.AppName.ToLower());
			return View(model);			
		}	

	}
}
