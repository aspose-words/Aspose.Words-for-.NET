using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class UnlockController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Unlock(string outputType, string passw)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				var docs =  UploadFiles(Request);

				AsposeWordsProtection asposeWordsProtection = new AsposeWordsProtection();
				response = asposeWordsProtection.Unlock(docs, outputType, passw);

			}

			return response;				
		}
		public ActionResult Unlock()
		{
			var model = new ViewModel(this, "Unlock")
			{
				ControlsView = "UnlockControls",
				SaveAsComponent = true,
				MaximumUploadFiles = 10,
				DropOrUploadFileLabel = Resources["DropOrUploadFiles"],
				ShowViewerButton = false
			};
			if (model.RedirectToMainApp)
				return Redirect("/words/" + model.AppName.ToLower());
			return View(model);			
		}	

	}
}
