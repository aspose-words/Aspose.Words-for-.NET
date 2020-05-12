using Aspose.Words.Live.Demos.UI.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Models
{
    public class FileUploadResult
    {
        public string LocalFilePath { get; set; }
        public string FileName { get; set; }
        public string FolderId { get; set; }
        public long FileLength { get; set; }

		public override string ToString()
		{
			
			return $"{200}|{FileName}|{FolderId}";
		}

	}

	public class FileUploadResponse
	{
		public string LocalFilePath { get; set; }
		public string FileName { get; set; }
		public string FolderId { get; set; }
		public long FileLength { get; set; }

		

	}
}
