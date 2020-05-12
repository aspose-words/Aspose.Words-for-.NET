using System;
using System.Collections.Generic;
using System.Text;
using System.Web;
using System.Xml;
using System.Web.Caching;
using System.IO;

namespace Aspose.Words.Live.Demos.UI.Config
{
    /// <summary>
    /// Acts as a helper class for the website's global.asax file
    /// </summary>
    public class GlobalAppHelper
    {

		public GlobalAppHelper(HttpContext hc, HttpApplicationState appState, string SessionID, string language)
		{
			
			AsposeAppContext.atcc = new AsposeAppContext(hc); // sync the context

			string ResourcesFile = hc.Server.MapPath("~/App_Data/resources_EN" + ".xml");   // reference all the extra info and resources files       
			if (language.Trim() != "")
			{
				string filPath = hc.Server.MapPath("~/App_Data/resources_" + language + ".xml");
				if (File.Exists(filPath))
				{
					ResourcesFile = filPath;
				}
			}

			// Load info from all these files into the cache
			initResources(ResourcesFile, SessionID);
		}
		/// <summary>
		/// Reads/parses the resources file and loads them into the cache in form of a dictionary
		/// </summary>
		/// <param name="ResourcesFile"></param>
		private void initResources(string ResourcesFile, string SessionID)
		{
			SessionID = "R" + SessionID;
			if (AsposeAppContext.Current.Cache[SessionID] == null)
			{
				// Added to solve the file not found problem, the wait is one time only when the application initializes or associated files are modified
				//System.Threading.Thread.Sleep(500);
				Dictionary<string, string> resources = new Dictionary<string, string>();
				XmlDocument xd = new XmlDocument();
				// TextWriter tr = (TextWriter)File.CreateText("F:\\assets\\my.log");tr.WriteLine(ResourcesFile);
				// tr.WriteLine(File.Exists(ResourcesFile)); tr.Close();
				if (ResourcesFile.Trim() != "")
				{
					xd.Load(ResourcesFile);
				}
				XmlNodeList xl = xd.SelectNodes("resources/res"); // use xpath to reach the res tag within resources
				foreach (XmlNode n in xl) // read the name attribute for key name and values from in between the tags
					resources.Add(n.Attributes["name"].Value, n.InnerText);

				// Add this dictionary into the cache with no expiration and associate a reload method in case of file change

				
				AsposeAppContext.Current.Cache.Add(
				   SessionID,
					resources,
					new CacheDependency(ResourcesFile),
					Cache.NoAbsoluteExpiration,
					Cache.NoSlidingExpiration,
					CacheItemPriority.NotRemovable,
					delegate (string key, object value, CacheItemRemovedReason reason) { initResources(ResourcesFile, SessionID); });
			}
		}


	}
}
