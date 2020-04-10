using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Web.Mvc;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	/// <summary>
	/// Common  API controller.
	/// </summary>
	public  class CommonController : BaseController
	{
		public override string Product => (string)RouteData.Values["product"];
		/// <summary>
		/// Sends back specified file from specified folder inside OutputDirectory.
		/// </summary>
		/// <param name="folder">Folder inside OutputDirectory.</param>
		/// <param name="file">File.</param>
		/// <returns>HTTP response with file.</returns>


		public FileResult DownloadFile(string fileName, string folderName)
		{
			var pathProcessor = new PathProcessor(folderName, file: fileName);
			
			return File(pathProcessor.DefaultOutFile, "application/octet-stream", fileName);
		}
		[HttpPost]
		public  FileUploadResult UploadFile()
		{
			FileUploadResult uploadResult = null;
			string fn = "";
			
			try
			{
				string _folderID = Guid.NewGuid().ToString();
				var pathProcessor = new PathProcessor(_folderID);
				if (Request.Files.Count > 0)
				{

					foreach (string fileName in Request.Files)
					{
						HttpPostedFileBase postedFile = Request.Files[fileName];
						fn = System.IO.Path.GetFileName(postedFile.FileName);
						if (postedFile != null)
						{
							// Check if File is available.
							if (postedFile != null && postedFile.ContentLength > 0)
							{
								postedFile.SaveAs(Path.Combine(pathProcessor.SourceFolder, fn));
							}
						}
					}
				}
				return new FileUploadResult
				{
					FileName = fn,
					FolderId = _folderID
				};
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			return uploadResult;
		}





	}
}
