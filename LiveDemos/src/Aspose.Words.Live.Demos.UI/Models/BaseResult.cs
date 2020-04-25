using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Models
{
	/// <summary>
	/// Base class for results.
	/// </summary>
	public class BaseResult
	{
		/// <summary>
		/// Is result success?
		/// </summary>
		public bool IsSuccess { get;  set; }

		/// <summary>
		/// idError.
		/// </summary>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "id always small")]
		public string idError { get;  set; }
	}
}
