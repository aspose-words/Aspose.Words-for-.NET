using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class ComparisonController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Comparison(string outputType)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs =  UploadDocuments(Request, _sourceFolder);

				AsposeWordsComparison wordsComparison = new AsposeWordsComparison();
				response = wordsComparison.Compare(docs, _sourceFolder);

			}

			return response;			
				
		}		

		public ActionResult Comparison()
		{
			var model = new ViewModel(this, "Comparison")
			{
				
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};
			
			if (model.RedirectToMainApp)
				return Redirect("/words/" + model.AppName.ToLower());
			return View(model);
			
		}

	}
}
