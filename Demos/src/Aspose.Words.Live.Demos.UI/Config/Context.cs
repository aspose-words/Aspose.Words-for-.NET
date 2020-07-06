using System;
using System.Collections.Generic;
using System.Text;
using System.Web;
using System.Web.Security;
using System.Web.Caching;
using System.Collections;
using System.Threading;

namespace Aspose.Words.Live.Demos.UI.Config
{
	/// <summary>
	/// Base class for all context objects used in active application
	/// </summary>
	public abstract class Context
	{
		protected HttpContext _context;

		/// <summary>
		/// Creates a customized context extending specified http context hc
		/// </summary>
		/// <param name="hc"></param>
		public Context(HttpContext hc)
		{
			_context = hc;

			//if (_context.applicationa != null)
			//{
			//  _host = _context.Request.Url.Host.ToLower();
			//}
			//else
		}

		/// <summary>
		/// simple cache wrapper
		/// </summary>
		public Cache Cache
		{
			get { return _context.Cache; }
		}
		/// <summary>
		/// simple session wrapper
		/// </summary>
		public System.Web.SessionState.HttpSessionState Session
		{
			get { return _context.Session; }
		}
		/// <summary>
		/// Stores the specified key value pair in the cache indefinitely, removed only on application reset or explicit removal
		/// </summary>
		/// <param name="key"></param>
		/// <param name="value"></param>
		public void PermanentAddtoCache(string key, object value)
		{
			_context.Cache.Insert(key, value, null, Cache.NoAbsoluteExpiration, Cache.NoSlidingExpiration, CacheItemPriority.NotRemovable, null);
		}

		protected string Locale =>
			_context.Request.Url.Host.StartsWith("zh.")
			? "ZH"
			: "EN";
		/// <summary>
		/// key/value based storage for all the error messages picked up from resources.xml file
		/// </summary>
		protected Dictionary<string, string> Resources
		{
			get
			{
				string sessionID = Configuration.ResourceFileSessionName;
				return (Dictionary<string, string>)Cache["R" + sessionID];
			}
		}
		/// <summary>
		/// Simple cookie wrapper
		/// </summary>
		public HttpCookieCollection Cookies
		{
			get { return _context.Request.Cookies; }
		}

		/// <summary>
		/// Checks if the session is valid i.e. not expired
		/// </summary>
		protected bool IsValid
		{
			get { return _context.Session != null; }
		}

		/// <summary>
		/// Simple Application wrapper
		/// </summary>
		private HttpApplicationState Application
		{
			get { return _context.Application; }
		}
	}
}
