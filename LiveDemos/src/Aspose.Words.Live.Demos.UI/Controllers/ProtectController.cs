using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class ProtectController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Protect(string passw)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				var docs =  UploadFiles(Request);

				AsposeWordsProtection asposeWordsProtection = new AsposeWordsProtection();
				response = asposeWordsProtection.Protect(docs, passw);
			}

			return response;				
		}
		public ActionResult Protect()
		{
			var model = new ViewModel(this, "Protect")
			{
				ControlsView = "UnlockControls",				
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"],
				ShowViewerButton = false
			};
			if (model.RedirectToMainApp)
				return Redirect("/words/" + model.AppName.ToLower());
			return View(model);			
		}	

	}
}
