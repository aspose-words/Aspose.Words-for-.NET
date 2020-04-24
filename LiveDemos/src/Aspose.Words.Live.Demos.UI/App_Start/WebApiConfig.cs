using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Net.Http.Headers;

namespace Aspose.Words.Live.Demos.UI
{
	///<Summary>
	/// WebApiConfig
	///</Summary>
	public static class WebApiConfig
	{
		///<Summary>
		/// Register
		///</Summary>
		public static void Register(HttpConfiguration config)
		{
			// Web API configuration and services
			// config.EnableCors();
			// Web API routes
			config.MapHttpAttributeRoutes();

			config.Routes.MapHttpRoute(
				name: "DefaultApi",
				routeTemplate: "api/{controller}/{action}/{id}",
				defaults: new { id = RouteParameter.Optional }
			);
			config.Formatters.JsonFormatter.SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/octet-stream"));
			//config.EnableCors();
		}
	}
}