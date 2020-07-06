using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class MergerController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Merger(string outputType)
		{
			Response response = null;
			
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs =  UploadDocuments(Request, _sourceFolder);

				AsposeWordsMerger wordsMerger = new AsposeWordsMerger();
				response = wordsMerger.Merge(docs, outputType, _sourceFolder);

			}

			return response;			
				
		}

		

		public ActionResult Merger()
		{
			var model = new ViewModel(this, "Merger")
			{
				SaveAsComponent = true,
				SaveAsOriginal = false,
				MaximumUploadFiles = 10,
				UseSorting = true,
				DropOrUploadFileLabel = Resources["DropOrUploadFiles"]
			};

			return View(model);
		}
		

	}
}
