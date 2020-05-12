using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;

namespace Aspose.Words.Live.Demos.UI.Helpers
{
	///<Summary>
	/// MultipartFormDataStreamProviderSafe  class
	///</Summary>
	public class MultipartFormDataStreamProviderSafe : MultipartFormDataStreamProvider
	{
		///<Summary>
		/// initialize MultipartFormDataStreamProviderSafe  class
		///</Summary>
		public MultipartFormDataStreamProviderSafe(string rootPath) : base(rootPath) { }
		///<Summary>
		/// GetLocalFileName method to get local file name from header
		///</Summary>
		public override string GetLocalFileName(HttpContentHeaders headers)
		{
			var fileName = headers?.ContentDisposition?.FileName;
			if (fileName != null)
			{
				fileName = fileName.TrimEnd('"').TrimStart('"');
				try
				{
					fileName = Path.GetFileName(fileName);
          if (!string.IsNullOrEmpty(fileName))
          {
            var name = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            if (System.IO.File.Exists(Path.Combine(RootPath, name + extension)))
            {
              var i = 2;
              while (System.IO.File.Exists(Path.Combine(RootPath, name + " " + i + extension)))
                i++;
              name += " " + i;
            }
            return name + extension;
          }
        }
				catch
				{
				}
			}

			return base.GetLocalFileName(headers);
		}
	}
}
