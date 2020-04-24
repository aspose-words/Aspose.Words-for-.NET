using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models;

using File = System.IO.File;


namespace Aspose.Words.Live.Demos.UI.Controllers
{
	// If the request is cross-origin, a browser makes the OPTIONS request at first
#if DEBUG
	//[System.Web.Http.Cors.EnableCors(origins: "http://localhost:2144", headers: "*", methods: "*")]
#endif

	///<Summary>
	/// AsposeAssemblyController
	///</Summary>
	public class AsposeWordsAssemblyController : ApiController
	{
		///<Summary>
		/// Upload
		///</Summary>
		[HttpPost]
		[MimeMultipart]
		[AcceptVerbs("GET", "POST")]
		public async Task<HttpResponseMessage> Upload(string folderName)
		{
			try
			{
				var provider = new MultipartFormDataStreamProvider(Config.Configuration.WorkingDirectory);
				await Request.Content.ReadAsMultipartAsync(provider);

				Directory.CreateDirectory(Path.Combine(Config.Configuration.WorkingDirectory, folderName));

				foreach (var file in provider.FileData)
				{
					var name = file.Headers.ContentDisposition.FileName.Trim('"');
					var path = Path.Combine(Config.Configuration.WorkingDirectory, folderName, name);
					File.Copy(file.LocalFileName, path, true);
					File.Delete(file.LocalFileName);
				}

				return Request.CreateResponse(HttpStatusCode.OK);
			}
			catch (Exception ex)
			{
				
				return Request.CreateResponse(HttpStatusCode.InternalServerError, ex);
			}
		}
		///<Summary>
		/// UploadCORS
		///</Summary>
		[HttpOptions]
		
		public HttpResponseMessage UploadCORS(string folderName)
		{
			return Request.CreateResponse(HttpStatusCode.OK);
		}
		///<Summary>
		/// Assemble
		///</Summary>
		[HttpPost]
		
		public Response Assemble(string productName, string folderName, string templateFilename, string datasourceFilename,
		  string datasourceName, int datasourceTableIndex = 0, string delimiter = ",")
		{
			AsposeWordsAssembly AsposeWordsAssembly = new AsposeWordsAssembly();
			return AsposeWordsAssembly.Assemble(folderName, templateFilename, datasourceFilename, datasourceName,
					  datasourceTableIndex, HttpUtility.UrlDecode(delimiter));
		

			
		}
		///<Summary>
		/// AssembleCORS
		///</Summary>
		[HttpOptions]		
		public HttpResponseMessage AssembleCORS(string productName, string folderName, string templateFilename, string datasourceFilename,
		  string datasourceName, int datasourceTableIndex = 0, string delimiter = ",")
		{
			return Request.CreateResponse(HttpStatusCode.OK);
		}
	}
}
