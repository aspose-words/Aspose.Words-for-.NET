using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class SignatureController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		//[HttpPost]
		//public Response Signature(string outputType, signatureType)
		//{
		//	Response response = null;
		//	if (Request.Files.Count > 0)
		//	{
		//		string _sourceFolder = Guid.NewGuid().ToString();
		//		var docs =  UploadDocuments(Request, _sourceFolder);

		//		AsposeWordsConversion wordsConversion = new AsposeWordsConversion();
		//		response = wordsConversion.ConvertFile(docs, outputType, _sourceFolder);

		//	}

		//	return response;			
				
		//}

		

		public ActionResult Signature()
		{
			var model = new ViewModel(this, "Signature")
			{
				ControlsView = "SignatureControls",
				SaveAsComponent = true,
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};
			if (model.RedirectToMainApp)
				return Redirect("/words/" + model.AppName.ToLower());
			return View(model);
			
		}
		

	}
}
