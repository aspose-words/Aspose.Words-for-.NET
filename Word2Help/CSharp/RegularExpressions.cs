//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//22/9/06 by Vladimir Averkin
using System.Text.RegularExpressions;

namespace Word2Help
{
    /// <summary>
    /// Central storage for regular expressions used in the project.
    /// </summary>
    public class RegularExpressions
    {
        // This class is static. No instance creation is allowed.
        private RegularExpressions() {}

        /// <summary>
        /// Regular expression specifying html title (framing tags excluded).
        /// </summary>
        public static Regex HtmlTitle
        {
            get 
            {
                if (gHtmlTitle == null) 
                {
                    gHtmlTitle = new Regex(HtmlTitlePattern,
                        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);
                }
                return gHtmlTitle;
            }
        }

        /// <summary>
        /// Regular expression specifying html head.
        /// </summary>
        public static Regex HtmlHead
        {
            get 
            {
                if (gHtmlHead == null) 
                {
                    gHtmlHead = new Regex(HtmlHeadPattern,
                        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);
                }
                return gHtmlHead;
            }
        }

        /// <summary>
        /// Regular expression specifying space right after div keyword in the first div declaration of html body.
        /// </summary>
        public static Regex HtmlBodyDivStart
        {
            get 
            {
                if (gHtmlBodyDivStart == null) 
                {
                    gHtmlBodyDivStart = new Regex(HtmlBodyDivStartPattern,
                        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);
                }
                return gHtmlBodyDivStart;
            }
        }

        private const string HtmlTitlePattern = @"(?<=\<title\>).*?(?=\</title\>)";
        private static Regex gHtmlTitle;

        private const string HtmlHeadPattern = @"\<head\>.*?\</head\>";
        private static Regex gHtmlHead;

        private const string HtmlBodyDivStartPattern = @"(?<=\<body\>\s*\<div)\s";
        private static Regex gHtmlBodyDivStart;
    }
}
