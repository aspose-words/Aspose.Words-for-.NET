using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using Aspose.Words.Live.Demos.UI.Config;
using Aspose.Words.Live.Demos.UI.Controllers;

namespace Aspose.Words.Live.Demos.UI.Models
{
	public class LandingPageModel
	{
		public BaseController Controller;
		/// <summary>
		/// Name of the product (e.g., words)
		/// </summary>
		public string Product { get; set; }
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

		public LandingPageModel(BaseController controller)
		{
			Controller = controller;
			Resources = controller.Resources;
			
		}
		private string GetFromResources(string key, string defaultKey = null)
		{
			if (Resources.ContainsKey(key))
				return Resources[key];
			if (!string.IsNullOrEmpty(defaultKey) && Resources.ContainsKey(defaultKey))
				return Resources[defaultKey];
			return "";
		}

		
	}
}
