//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

namespace Word2HelpExample
{
    /// <summary>
    /// Represents a single topic that will be written as an HTML file.
    /// </summary>
    public class Topic
    {
        /// <summary>
        /// Creates a topic.
        /// </summary>
        public Topic(Section section, string fixUrl)
        {
            mTopicDoc = new Document();
            mTopicDoc.AppendChild(mTopicDoc.ImportNode(section, true, ImportFormatMode.KeepSourceFormatting));
            mTopicDoc.FirstSection.Remove();

            Paragraph headingPara = (Paragraph)mTopicDoc.FirstSection.Body.FirstChild;
            if (headingPara == null)
                ThrowTopicException("The section does not start with a paragraph.", section);

            mHeadingLevel = headingPara.ParagraphFormat.StyleIdentifier - StyleIdentifier.Heading1;
            if ((mHeadingLevel < 0) || (mHeadingLevel > 8))
                ThrowTopicException("This topic does not start with a heading style paragraph.", section);
            
            mTitle = headingPara.GetText().Trim();
            if (mTitle == "")
                ThrowTopicException("This topic heading does not have text.", section);

            // We actually remove the heading paragraph because <h1> will be output in the banner.
            headingPara.Remove();

            mTopicDoc.BuiltInDocumentProperties.Title = mTitle;

            FixHyperlinks(section.Document, fixUrl);
        }

        private static void ThrowTopicException(string message, Section section)
        {
            throw new Exception(message + " Section text: " + section.Body.ToString(SaveFormat.Text).Substring(0, 50));
        }

        private void FixHyperlinks(DocumentBase originalDoc, string fixUrl)
        {
            if (fixUrl.EndsWith("/"))
                fixUrl = fixUrl.Substring(0, fixUrl.Length - 1);

            NodeCollection fieldStarts = mTopicDoc.GetChildNodes(NodeType.FieldStart, true);
            foreach (FieldStart fieldStart in fieldStarts)
            {
                if (fieldStart.FieldType != FieldType.FieldHyperlink)
                    continue;

                Hyperlink hyperlink = new Hyperlink(fieldStart);
                if (hyperlink.IsLocal)
                {
                    // We use "Hyperlink to a place in this document" feature of Microsoft Word
                    // to create local hyperlinks between topics within the same doc file.
                    // It causes MS Word to auto generate the bookmark name.
                    string bmkName = hyperlink.Target;

                    // But we have to follow the bookmark to get the text of the topic heading paragraph
                    // in order to be able to build the proper filename of the topic file.
                    Bookmark bmk = originalDoc.Range.Bookmarks[bmkName];

                    if (bmk == null)
                        throw new Exception(string.Format("Found a link to a bookmark, but cannot locate the bookmark. Name:'{0}'.", bmkName));

                    Paragraph para = (Paragraph)bmk.BookmarkStart.ParentNode;
                    string topicName = para.GetText().Trim();

                    hyperlink.Target = HeadingToFileName(topicName) + ".html";
                    hyperlink.IsLocal = false;
                }
                else
                {
                    // We "fix" URL like this:
                    // http://www.aspose.com/Products/Aspose.Words/Api/Aspose.Words.Body.html
                    // by changing them into this:
                    // Aspose.Words.Body.html
                    if (hyperlink.Target.StartsWith(fixUrl) &&
                        (hyperlink.Target.Length > (fixUrl.Length + 1)))
                    {
                        hyperlink.Target = hyperlink.Target.Substring(fixUrl.Length + 1);
                    }
                }
            }
        }

        public void WriteHtml(string htmlHeader, string htmlBanner, string htmlFooter, string outDir)
        {
            string fileName = Path.Combine(outDir, FileName);

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.PrettyFormat = true;
            // This is to allow headings to appear to the left of main text.
            saveOptions.AllowNegativeLeftIndent = true;
            // Disable headers and footers.
            saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.None;

            // Export the document to HTML.
            mTopicDoc.Save(fileName, saveOptions);

            // We need to modify the HTML string, read HTML back.
            string html;
            using (StreamReader reader = new StreamReader(fileName))
                html = reader.ReadToEnd();

            // Builds the HTML <head> element.
            string header = RegularExpressions.HtmlTitle.Replace(htmlHeader, mTitle, 1);
            
            // Applies the new <head> element instead of the original one.
            html = RegularExpressions.HtmlHead.Replace(html, header, 1);
            html = RegularExpressions.HtmlBodyDivStart.Replace(html, @" id=""nstext""", 1);

            string banner = htmlBanner.Replace("###TOPIC_NAME###", mTitle);
            
            // Add the standard banner.
            html = html.Replace("<body>", "<body>" + banner);
            
            // Add the standard footer.
            html = html.Replace("</body>", htmlFooter + "</body>");

            using (StreamWriter writer = new StreamWriter(fileName))
                writer.Write(html);
        }

        /// <summary>
        /// Removes various characters from the heading to form a file name that does not require escaping.
        /// </summary>
        private static string HeadingToFileName(string heading)
        {
            StringBuilder b = new StringBuilder();
            foreach (char c in heading)
            {
                if (Char.IsLetterOrDigit(c))
                    b.Append(c);
            }

            return b.ToString();
        }

        public Document Document
        {
            get { return mTopicDoc; }
        }

        /// <summary>
        /// Gets the name of the topic html file without path.
        /// </summary>
        public string FileName
        {
            get { return HeadingToFileName(mTitle) + ".html"; }
        }

        public string Title
        {
            get { return mTitle; }
        }

        public int HeadingLevel
        {
            get { return mHeadingLevel; }
        }

        /// <summary>
        /// Returns true if the topic has no text (the heading paragraph has already been removed from the topic).
        /// </summary>
        public bool IsHeadingOnly
        {
            get
            {
                Body body = mTopicDoc.FirstSection.Body;
                return (body.FirstParagraph == null);
            }
        }

        private readonly Document mTopicDoc;
        private readonly string mTitle;
        private readonly int mHeadingLevel;
    }
}