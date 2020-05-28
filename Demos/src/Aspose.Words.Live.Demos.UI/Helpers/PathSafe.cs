using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Helpers
{
	/// <summary>
	/// PathSafe Class
	/// </summary>
	public class PathSafe
	{
		/// <summary>
		/// Method to generate path based on base directory and parts.
		/// Also this method checks and prevent Path Traversal vulnerability.
		/// </summary>
		/// <param name="baseDir">Base path. Path Traversal checked against this base. All result path must be child for this base.</param>
		/// <param name="parts">Path parts.</param>
		/// <returns>Path.</returns>
		public static string GenerateAndValidate(string baseDir, params string[] parts)
		{
			baseDir = Path.GetFullPath(baseDir);
			var partsList = parts.ToList();
			partsList.Insert(0, baseDir);
			var path = Path.GetFullPath(Path.Combine(partsList.ToArray()));
			if (!path.StartsWith(baseDir))
				HttpHelper.Throw400(null);

			return path;
		}
	}
}
