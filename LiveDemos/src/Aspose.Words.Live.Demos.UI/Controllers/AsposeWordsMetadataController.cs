using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words;
using Aspose.Words.Properties;
using Newtonsoft.Json.Linq;


namespace Aspose.Words.Live.Demos.UI.Controllers
{
	///<Summary>
	/// AsposeWordsMetadataController class to parse word document
	///</Summary>
	public class AsposeWordsMetadataController : AsposeWordsBase
	{
    ///<Summary>
    /// Properties method. Should include 'FileName' and 'id' as params
    ///</Summary>
    [HttpPost]
		
		public HttpResponseMessage Properties(string folderName, string fileName)
    {
      
			

	  try
      {
        

        var doc = new Document(Path.Combine(Config.Configuration.WorkingDirectory , folderName, fileName));
        //doc.UpdateWordCount();
        return Request.CreateResponse(HttpStatusCode.OK, new PropertiesResponse(doc));
      }
      catch (Exception ex)
      {
				Console.WriteLine(ex.Message);
        return Request.CreateResponse(HttpStatusCode.InternalServerError, ex.Message);
      }
    }

    ///<Summary>
    /// Properties method. Should include 'FileName', 'id', 'properties' as params
    ///</Summary>
    [HttpPost]
		[AcceptVerbs("GET", "POST")]
		public Response Download()
    {
      Opts.AppName = "Metadata";
      Opts.MethodName = "Download";
      try
      {
        var request = Request.Content.ReadAsAsync<JObject>().Result;
				Opts.FileName = Convert.ToString(request["FileName"]);
				Opts.ResultFileName = Opts.FileName;
				Opts.FolderName = Convert.ToString(request["id"]);

				var doc = new Document(Opts.WorkingFileName);
				var pars = request["properties"]["BuiltIn"].ToObject<List<DocProperty>>();
        SetBuiltInProperties(doc, pars);
        pars = request["properties"]["Custom"].ToObject<List<DocProperty>>();
        SetCustomProperties(doc, pars);
				
				return  Process((inFilePath, outPath, zipOutFolder) => { doc.Save(outPath); });
				

	  }
      catch (Exception ex)
      {
				Console.WriteLine(ex.Message);
        return new Response
        {
          Status = "500 " + ex.Message,
          StatusCode = 500
        };
      }
    }

    ///<Summary>
    /// Properties method. Should include 'FileName', 'id' as params
    ///</Summary>
    [HttpPost]
		[AcceptVerbs("GET", "POST")]
		public Response Clear()
    {
      Opts.AppName = "Metadata";
      Opts.MethodName = "Clear";
      try
      {
        var request = Request.Content.ReadAsAsync<JObject>().Result;
        Opts.FileName = Convert.ToString(request["FileName"]);
        Opts.ResultFileName = Opts.FileName;
        Opts.FolderName = Convert.ToString(request["id"]);

        var doc = new Document(Opts.WorkingFileName);
        doc.BuiltInDocumentProperties.Clear();
        doc.CustomDocumentProperties.Clear();

        return  Process((inFilePath, outPath, zipOutFolder) => { doc.Save(outPath); });
      }
      catch (Exception ex)
      {
				Console.WriteLine(ex.Message);
        return new Response
        {
          Status = "500 " + ex.Message,
          StatusCode = 500
        };
      }
    }
    
    /// <summary>
    /// SetBuiltInProperties
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="pars"></param>
    private void SetBuiltInProperties(Document doc, List<DocProperty> pars)
    {
      var builtin = doc.BuiltInDocumentProperties;
      var t = builtin.GetType();
      foreach (var par in pars)
      {
        var prop = t.GetProperty(par.Name);
        if (prop != null)
          switch (par.Type)
          {
            case PropertyType.String:
              prop.SetValue(builtin, Convert.ToString(par.Value));
              break;
            case PropertyType.Boolean:
              prop.SetValue(builtin, Convert.ToBoolean(par.Value));
              break;
            case PropertyType.Number:
              prop.SetValue(builtin, Convert.ToInt32(par.Value));
              break;
            case PropertyType.DateTime:
              prop.SetValue(builtin, Convert.ToDateTime(par.Value));
              break;
            case PropertyType.Double:
              prop.SetValue(builtin, Convert.ToDouble(par.Value));
              break;
          }
      }
    }
    

    /// <summary>
    /// SetCustomProperties
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="pars"></param>
    private void SetCustomProperties(Document doc, List<DocProperty> pars)
    {
      var custom = doc.CustomDocumentProperties;
      custom.Clear();
      foreach (var par in pars)
        switch (par.Type)
        {
          case PropertyType.String:
            custom.Add(par.Name, Convert.ToString(par.Value));
            break;
          case PropertyType.Boolean:
            custom.Add(par.Name, Convert.ToBoolean(par.Value));
            break;
          case PropertyType.Number:
            custom.Add(par.Name, Convert.ToInt32(par.Value));
            break;
          case PropertyType.DateTime:
            custom.Add(par.Name, Convert.ToDateTime(par.Value));
            break;
          case PropertyType.Double:
            custom.Add(par.Name, Convert.ToDouble(par.Value));
            break;
        }
    }

    /// <summary>
    /// PropertiesResponse
    /// </summary>
    private class PropertiesResponse
    {
      public BuiltInDocumentProperties BuiltIn { get; set; }
      public CustomDocumentProperties Custom { get; set; }

      public PropertiesResponse(Document doc)
      {
        BuiltIn = doc.BuiltInDocumentProperties;
        Custom = doc.CustomDocumentProperties;
      }
    }

    /// <summary>
    /// The same fields as in DocumentProperty
    /// </summary>
    private class DocProperty
    {
      public string Name { get; set; }
      public object Value { get; set; }
      public PropertyType Type { get; set; }
    }
  }
}
