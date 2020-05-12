using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// License class to set apose products license
	///</Summary>
	public static class License
	{
		private static string _licenseFileName = "Aspose.Total.lic";

			
		///<Summary>
		/// SetAsposeWordsLicense method to Aspose.Words License
		///</Summary>
		public static void SetAsposeWordsLicense()
		{
			try
			{
				Aspose.Words.License awLic = new Aspose.Words.License();
				awLic.SetLicense(_licenseFileName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
		
		
	}
}
