using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Models
{
	/// <summary>
	/// File processing result.
	/// </summary>
	public class FileSafeResult : BaseResult
	{
		/// <summary>
		/// Upload id.
		/// </summary>
		public string id { get; set; }
		/// <summary>
		/// File name.
		/// </summary>
		public string FileName { get; set; }

		/// <summary>
		/// idUpload from request.		
		/// </summary>
		public string idUpload { get; set; }

		/// <summary>
		/// File length.
		/// </summary>
		public long? FileLength => _localFilePath == null ? (long?)null : new FileInfo(_localFilePath).Length;

		/// <summary>
		/// Used to stores local file path.
		/// </summary>
		protected string _localFilePath;
		/// <summary>
		/// Returns local file path.
		/// </summary>
		/// <returns>Local file path.</returns>
		public string GetLocalFilePath() => _localFilePath;

		/// <summary>
		/// FileSafeResult constructor.
		/// Sets IsSuccess to true.
		/// </summary>
		public FileSafeResult()
		{
			this.IsSuccess = true;
		}

		/// <summary>
		/// Internal FileSafeResult constructor.
		/// Used to set local file path.
		/// </summary>
		/// <param name="localFilePath"></param>
		internal FileSafeResult(string localFilePath) : this()
		{
			_localFilePath = localFilePath;
		}
	}
}
