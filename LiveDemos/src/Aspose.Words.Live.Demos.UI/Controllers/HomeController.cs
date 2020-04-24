using Aspose.Words.Live.Demos.UI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	public class HomeController : BaseController
	{
	
		public override string Product => (string)RouteData.Values["productname"];
		

		public ActionResult Index()
		{
			ViewBag.PageTitle = "Free C# MVC Word Document Processing APPs - aspose.app";
			ViewBag.MetaDescription = "100% free apps for DOC, DOCX, DOT, DOTX, RTF, ODT, OTT, TXT, HTML, XHTML, MHTML files. View, convert, split, compare, sign, watermark, merge or redact content from Word Processing files.";
			var model = new LandingPageModel(this)

			{
				Product = Product
			};

			return View(model);
		}		

		public ActionResult Default()
		{
			ViewBag.PageTitle = "Free C# MVC Word Document Processing APPs - aspose.app";
			ViewBag.MetaDescription = "100% free apps for DOC, DOCX, DOT, DOTX, RTF, ODT, OTT, TXT, HTML, XHTML, MHTML files. View, convert, split, compare, sign, watermark, merge or redact content from Word Processing files.";
			var model = new LandingPageModel(this);

			return View(model);
		}
	}
}
