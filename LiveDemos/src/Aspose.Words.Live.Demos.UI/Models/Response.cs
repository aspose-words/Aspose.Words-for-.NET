using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Web;
using System.Collections.ObjectModel;

namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// Response class to get or set any api call status
	///</Summary>
	public class Response
	{
		///<Summary>
		/// Get or set DownloadFileLink
		///</Summary>
		public string DownloadFileLink { get; set; }
		///<Summary>
		/// Get or set StatusCode
		///</Summary>
		public int StatusCode { get; set; }
		///<Summary>
		/// Get or set FileName
		///</Summary>
		public string FileName { get; set; }
		///<Summary>
		/// Get or set FolderName
		///</Summary>
		public string FolderName { get; set; }
		///<Summary>
		/// Get or set Status
		///</Summary>
		public string Status { get; set; }
		///<Summary>
		/// Get or set Text
		///</Summary>
		public string Text { get; set; }
		///<Summary>
		/// Get or set Files
		///</Summary>
		public Collection<string> Files;
		///<Summary>
		/// Get or set FileProcessingErrorCode
		///</Summary>
		public FileProcessingErrorCode FileProcessingErrorCode { get; set; }
		public string DownloadURL()
		{
			var url = new StringBuilder( Config.Configuration.FileDownloadLink);
			url.Append("?FileName=");
			url.Append(HttpUtility.UrlPathEncode(FileName));			
			if (!string.IsNullOrEmpty(FolderName))
			{
				url.Append("&FolderName=");
				url.Append(FolderName);
			}

			return url.ToString();
		}
		public string ViewerURL(string product, string callbackURL)
		{
			var url = new StringBuilder();
			switch (product.ToLower())
			{
				case "words":
					url.Append($"/{product}/view?FolderName=");
					url.Append(FolderName);
					url.Append("&FileName=");
					url.Append(HttpUtility.UrlEncode(FileName));
					url.Append("&callbackURL=");
					url.Append(callbackURL);
					break;
				
				default:
					//url.Append(string.Format(Configuration.FileViewLink, product));
					url.Append(string.Format("", product));
					url.Append("/");
					url.Append(FolderName);
					url.Append("?fileName=");
					url.Append(HttpUtility.UrlEncode(FileName));
					url.Append("&callbackURL=");
					url.Append(callbackURL);
					break;
			}
			return url.ToString();
		}

		/// <summary>
		/// Generate the viewer url with the product name by the file extension
		/// </summary>
		/// <param name="callbackURL"></param>
		/// <returns></returns>
		public string ViewerURL(string callbackURL)
		{
			return ViewerURL(GetProductByFileExtension(), callbackURL);
		}

		private string GetProductByFileExtension()
		{
			switch (Path.GetExtension(FileName)?.ToLower())
			{
				case ".eml":
				case ".msg":
					return "email";
				case ".mhtml":
				case ".html":
					return "html";
				case ".epub":
				case ".ps":
				case ".xps":
				case ".pdf":
					return "pdf";
				default:
					return "words";
			}
		}

		public override string ToString()
		{
			//return $"{StatusCode} - {Status}";
			//return $"{StatusCode}|{HttpUtility.UrlPathEncode(FileName)}|{FolderName}";

			int _fileProcessingErrorCode = (int)Enum.Parse(typeof(Aspose.Words.Live.Demos.UI.Models.FileProcessingErrorCode), FileProcessingErrorCode.ToString());
			return $"{StatusCode}|{FileName}|{FolderName}|{_fileProcessingErrorCode}";
		}

	}
}
