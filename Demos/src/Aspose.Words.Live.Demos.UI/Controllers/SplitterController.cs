using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class SplitterController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Splitter(string outputType, string splitType, string pars)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs =  UploadDocuments(Request, _sourceFolder);

				AsposeWordsSplitter wordsSplitter = new AsposeWordsSplitter();
				response = wordsSplitter.Split(docs, _sourceFolder, outputType, int.Parse( splitType), pars);

			}

			return response;			
				
		}

		

		public ActionResult Splitter()
		{
			

			var model = new ViewModel(this, "Splitter")
			{
				ControlsView = "SplitterControls",
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
