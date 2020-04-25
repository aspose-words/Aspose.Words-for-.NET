using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Helpers
{
	/// <summary>
	/// Extenders
	/// </summary>
	internal static class Extenders
	{
		/// <summary>
		/// Parses string as enum of specified type.
		/// </summary>
		/// <typeparam name="T">Enum type.</typeparam>
		/// <param name="source">Source string.</param>
		/// <param name="ignoreCase">Is operation case insensitive.</param>
		/// <returns>Enum value.</returns>
		public static T ParseEnum<T>(this string source, bool ignoreCase = true) where T : struct => (T)Enum.Parse(typeof(T), source, ignoreCase);
	}
}
