// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//14/9/06 by Vladimir Averkin
using System;
using System.IO;
using System.Collections;
using System.Text;
using System.Xml;
using Aspose.Words;

namespace Word2Help
{
    /// <summary>
    /// This is the main class.
    /// Loads Word document(s), splits them into topics, saves HTML files and builds content.xml.
    /// </summary>
    public class TopicCollection
    {
        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="htmlTemplatesDir">The directory that contains header.html, banner.html and footer.html files.</param>
        /// <param name="fixUrl">The url that will be removed from any hyperlinks that start with this url.
        /// This allows turning some absolute URLS into relative.</param>
        public TopicCollection(string htmlTemplatesDir, string fixUrl) 
        {
            mTopics = new ArrayList();
            mFixUrl = fixUrl;
            mHtmlHeader = ReadFile(Path.Combine(htmlTemplatesDir, "header.html")); 
            mHtmlBanner = ReadFile(Path.Combine(htmlTemplatesDir, "banner.html")); 
            mHtmlFooter = ReadFile(Path.Combine(htmlTemplatesDir, "footer.html")); 
        }

        /// <summary>
        /// Processes all DOC files found in the specified directory.
        /// Loads and splits each document into separate topics.
        /// </summary>
        public void AddFromDir(string dirName)
        {
            foreach (string filename in Directory.GetFiles(dirName, "*.doc"))
                AddFromFile(filename);
        }

        /// <summary>
        /// Processes a specified DOC file. Loads and splits into topics.
        /// </summary>
        public void AddFromFile(string fileName)
        {
            Document doc = new Document(fileName);
            InsertTopicSections(doc);
            AddTopics(doc);
        }

        /// <summary>
        /// Saves all topics as HTML files.
        /// </summary>
        public void WriteHtml(string outDir)
        {
            foreach (Topic topic in mTopics)
            {
                if (!topic.IsHeadingOnly)
                    topic.WriteHtml(mHtmlHeader, mHtmlBanner, mHtmlFooter, outDir);
            }
        }

        /// <summary>
        /// Saves the content.xml file that describes the tree of topics.
        /// </summary>
        public void WriteContentXml(string outDir)
        {
            XmlTextWriter writer = new XmlTextWriter(Path.Combine(outDir, "content.xml"), Encoding.UTF8);
            writer.Namespaces = false;
            writer.Formatting = Formatting.Indented;

            writer.WriteStartDocument(true);
            writer.WriteStartElement("content");
            writer.WriteAttributeString("dir", outDir);

            for (int i = 0; i < mTopics.Count; i++)
            {
                Topic topic = (Topic)mTopics[i];

                int nextTopicIdx = i + 1;
                Topic nextTopic = (nextTopicIdx < mTopics.Count) ? (Topic)mTopics[i + 1] : null;

                int nextHeadingLevel = (nextTopic != null) ? nextTopic.HeadingLevel : 0;

                if (nextHeadingLevel > topic.HeadingLevel)
                {
                    // Next topic is nested, therefore we have to start a book. 
                    // We only allow increase level at a time.
                    if (nextHeadingLevel != topic.HeadingLevel + 1)
                        throw new Exception("Topic is nested for more than one level at a time. Title: " + topic.Title);

                    WriteBookStart(writer, topic);
                }
                else if (nextHeadingLevel < topic.HeadingLevel)
                {
                    // Next topic is one or more levels higher in the outline.

                    // Write out the current topic.
                    WriteItem(writer, topic.Title, topic.FileName);

                    // End one or more nested topics could have ended at this point.
                    int levelsToClose = topic.HeadingLevel - nextHeadingLevel;
                    while (levelsToClose > 0)
                    {
                        WriteBookEnd(writer);
                        levelsToClose--;
                    }
                }
                else
                {
                    // A topic at the current level and it has no children.
                    WriteItem(writer, topic.Title, topic.FileName);
                }
            }

            writer.WriteEndElement();	// content
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }
        
        /// <summary>
        /// Inserts section breaks that delimit the topics.
        /// </summary>
        /// <param name="doc">The document where to insert the section breaks.</param>
        private static void InsertTopicSections(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true, false);
            ArrayList topicStartParas = new ArrayList();

            foreach (Paragraph para in paras)
            {
                StyleIdentifier style = para.ParagraphFormat.StyleIdentifier;
                if ((style >= StyleIdentifier.Heading1) && (style <= MaxTopicHeading) &&
                    (para.HasChildNodes))
                {
                    // Select heading paragraphs that must become topic starts.
                    // We can't modify them in this loop, we have to remember them in an array first.
                    topicStartParas.Add(para);
                }
                else if ((style > MaxTopicHeading) && (style <= StyleIdentifier.Heading9))
                {
                    // Pull up headings. For example: if Heading 1-4 become topics, then I want Headings 5+ 
                    // to become Headings 4+. Maybe I want to pull up even higher?
                    para.ParagraphFormat.StyleIdentifier = (StyleIdentifier)((int)style - 1);
                }
            }

            foreach (Paragraph para in topicStartParas)
            {
                Section section = para.ParentSection;

                // Insert section break if the paragraph is not at the beginning of a section already.
                if (para != section.Body.FirstParagraph)
                {
                    builder.MoveTo(para.FirstChild);
                    builder.InsertBreak(BreakType.SectionBreakNewPage);

                    // This is the paragraph that was inserted at the end of the now old section.
                    // We don't really need the extra paragraph, we just needed the section.
                    section.Body.LastParagraph.Remove();
                }
            }
        }

        /// <summary>
        /// Goes through the sections in the document and adds them as topics to the collection.
        /// </summary>
        private void AddTopics(Document doc)
        {
            foreach (Section section in doc.Sections)
            {
                try
                {
                    Topic topic = new Topic(section, mFixUrl);
                    mTopics.Add(topic);
                }
                catch (Exception e)
                {
                    // If one topic fails, we continue with others.
                    Console.WriteLine(e.Message);
                }
            }
        }

        private static void WriteBookStart(XmlWriter writer, Topic topic)
        {
            writer.WriteStartElement("book");
            writer.WriteAttributeString("name", topic.Title);

            if (!topic.IsHeadingOnly)
                writer.WriteAttributeString("href", topic.FileName);
        }

        private static void WriteBookEnd(XmlWriter writer)
        {
            writer.WriteEndElement();	// book
        }
        
        private static void WriteItem(XmlWriter writer, string name, string href)
        {
            writer.WriteStartElement("item");
            writer.WriteAttributeString("name", name);
            writer.WriteAttributeString("href", href);
            writer.WriteEndElement();	// item
        }

        private static string ReadFile(string fileName)
        {
            using (StreamReader reader = new StreamReader(fileName))
                return reader.ReadToEnd();
        }

        private readonly ArrayList mTopics;
        private readonly string mFixUrl;
        private readonly string mHtmlHeader;
        private readonly string mHtmlBanner;
        private readonly string mHtmlFooter;

        /// <summary>
        /// Specifies the maximum Heading X number. 
        /// All of the headings above or equal to this will be put into their own topics.
        /// </summary>
        private const StyleIdentifier MaxTopicHeading = StyleIdentifier.Heading4;
    }
}
