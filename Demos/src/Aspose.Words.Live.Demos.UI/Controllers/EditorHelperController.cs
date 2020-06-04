using System.IO;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	///<Summary>
	/// EditorHelperController class to call editor controller based 
	///</Summary>
	public class EditorHelperController : AsposeWordsBase
	{
        ///<Summary>
        /// GetHTML method to call GetHTML method based on product name
        ///</Summary>
        [HttpPost]
		[AcceptVerbs("GET", "POST")]
		public HttpResponseMessage GetHTML(string fileName, string folderName)
        {
			fileName = HttpUtility.UrlDecode(fileName);
          
                        var wordsController = new AsposeWordsEditorController();
                        var html = wordsController.GetHTML(fileName,folderName);
                        if (html != null)
                            return Request.CreateResponse(HttpStatusCode.OK, html);
                        return Request.CreateResponse(HttpStatusCode.InternalServerError, "Internal Server Error");
                   
        }
        ///<Summary>
        /// GetHTMLCORS method
        ///</Summary>
        [HttpOptions]        
        public HttpResponseMessage GetHTMLCORS()
        {
            return Request.CreateResponse(HttpStatusCode.OK);
        }
        ///<Summary>
        /// UpdateContentsRequest class to get or set UpdateContentsRequest properties
        ///</Summary>
        public class UpdateContentsRequest
        {
            ///<Summary>
            /// get or set filename
            ///</Summary>
            public string fileName { get; set; }

            ///<Summary>
            /// get or set folderName
            ///</Summary>
            public string folderName { get; set; }
            
            ///<Summary>
            /// get or set productName
            ///</Summary>
            public string productName { get; set; }
            ///<Summary>
            /// get or set htmldata
            ///</Summary>
            public string htmldata { get; set; }
            ///<Summary>
            /// get or set outputType
            ///</Summary>
            public string outputType { get; set; }
        }
		///<Summary>
		/// UpdateContents method to update contents 
		///</Summary>
		[HttpPost]
		[AcceptVerbs("GET", "POST")]
		public HttpResponseMessage UpdateContents([FromBody] UpdateContentsRequest request)
		{
			if (string.IsNullOrEmpty(request.outputType))
				request.outputType = Path.GetExtension(request.fileName);

			
						var wordsController = new AsposeWordsEditorController();
						var response = wordsController.UpdateContents(request.fileName, request.htmldata, request.outputType);
						if (response != null)
							return Request.CreateResponse(HttpStatusCode.OK, response);
						return Request.CreateResponse(HttpStatusCode.InternalServerError, "Internal Server Error");
				
			
        }
        ///<Summary>
        /// UpdateContents method to update contents
        ///</Summary>
        [HttpOptions]        
        public HttpResponseMessage UpdateContentsCORS()
        {
            return Request.CreateResponse(HttpStatusCode.OK);
        }
    }
}
