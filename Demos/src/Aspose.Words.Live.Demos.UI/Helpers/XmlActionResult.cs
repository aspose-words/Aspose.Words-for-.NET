using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Aspose.Words.Live.Demos.UI.Helpers
{
	public class XmlActionResult<T> : ActionResult
	{
		public XmlActionResult(T data, string @namespace = null)
		{
			Data = data;
			Namespace = @namespace;
		}

		public T Data { get; private set; }
		public string Namespace { get; private set; }

		public override void ExecuteResult(ControllerContext context)
		{
			context.HttpContext.Response.ContentType = "text/xml";

			var settings = new XmlWriterSettings
			{
				Encoding = Encoding.UTF8
			};

			using (XmlWriter xmlWriter = XmlWriter.Create(context.HttpContext.Response.OutputStream, settings))
			{
				if (Namespace == null)
				{
					new XmlSerializer(typeof(T)).Serialize(xmlWriter, Data);
				}
				else
				{
					var ns = new XmlSerializerNamespaces();
					ns.Add("", Namespace);

					new XmlSerializer(typeof(T), Namespace).Serialize(xmlWriter, Data, ns);
				}
			}
		}
	}
}
