using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Split_Documents
{
    internal class SplitIntoHtmlPages : DocsExamplesBase
    {
        [Test]
        public void HtmlPages()
        {
            string srcFileName = MyDir + "Footnotes and endnotes.docx";
            string tocTemplate = MyDir + "Table of content template.docx";

            string outDir = Path.Combine(ArtifactsDir, "HtmlPages");
            Directory.CreateDirectory(outDir);

            WordToHtmlConverter w = new WordToHtmlConverter();
            w.Execute(srcFileName, tocTemplate, outDir);
        }
    }

    internal class WordToHtmlConverter
    {
        /// <summary>
        /// Performs the Word to HTML conversion.
        /// </summary>
        /// <param name="srcFileName">The MS Word file to convert.</param>
        /// <param name="tocTemplate">An MS Word file that is used as a template to build a table of contents.
        /// This file needs to have a mail merge region called "TOC" defined and one mail merge field called "TocEntry".</param>
        /// <param name="dstDir">The output directory where to write HTML files.</param>
        internal void Execute(string srcFileName, string tocTemplate, string dstDir)
        {
            mDoc = new Document(srcFileName);
            mTocTemplate = tocTemplate;
            mDstDir = dstDir;

            List<Paragraph> topicStartParas = SelectTopicStarts();
            InsertSectionBreaks(topicStartParas);
            List<Topic> topics = SaveHtmlTopics();
            SaveTableOfContents(topics);
        }

        /// <summary>
        /// Selects heading paragraphs that must become topic starts.
        /// We can't modify them in this loop, so we need to remember them in an array first.
        /// </summary>
        private List<Paragraph> SelectTopicStarts()
        {
            NodeCollection paras = mDoc.GetChildNodes(NodeType.Paragraph, true);
            List<Paragraph> topicStartParas = new List<Paragraph>();

            foreach (Paragraph para in paras)
            {
                StyleIdentifier style = para.ParagraphFormat.StyleIdentifier;
                if (style == StyleIdentifier.Heading1)
                    topicStartParas.Add(para);
            }

            return topicStartParas;
        }

        /// <summary>
        /// Insert section breaks before the specified paragraphs.
        /// </summary>
        private void InsertSectionBreaks(List<Paragraph> topicStartParas)
        {
            DocumentBuilder builder = new DocumentBuilder(mDoc);
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
        /// Splits the current document into one topic per section and saves each topic
        /// as an HTML file. Returns a collection of Topic objects.
        /// </summary>
        private List<Topic> SaveHtmlTopics()
        {
            List<Topic> topics = new List<Topic>();
            for (int sectionIdx = 0; sectionIdx < mDoc.Sections.Count; sectionIdx++)
            {
                Section section = mDoc.Sections[sectionIdx];

                string paraText = section.Body.FirstParagraph.GetText();

                // Use the text of the heading paragraph to generate the HTML file name.
                string fileName = MakeTopicFileName(paraText);
                if (fileName == "")
                    fileName = "UNTITLED SECTION " + sectionIdx;

                fileName = Path.Combine(mDstDir, fileName + ".html");

                // Use the text of the heading paragraph to generate the title for the TOC.
                string title = MakeTopicTitle(paraText);
                if (title == "")
                    title = "UNTITLED SECTION " + sectionIdx;

                Topic topic = new Topic(title, fileName);
                topics.Add(topic);

                SaveHtmlTopic(section, topic);
            }

            return topics;
        }

        /// <summary>
        /// Leaves alphanumeric characters, replaces white space with underscore
        /// And removes all other characters from a string.
        /// </summary>
        private string MakeTopicFileName(string paraText)
        {
            StringBuilder b = new StringBuilder();
            foreach (char c in paraText)
            {
                if (char.IsLetterOrDigit(c))
                    b.Append(c);
                else if (c == ' ')
                    b.Append('_');
            }

            return b.ToString();
        }

        /// <summary>
        /// Removes the last character (which is a paragraph break character from the given string).
        /// </summary>
        private string MakeTopicTitle(string paraText)
        {
            return paraText.Substring(0, paraText.Length - 1);
        }

        /// <summary>
        /// Saves one section of a document as an HTML file.
        /// Any embedded images are saved as separate files in the same folder as the HTML file.
        /// </summary>
        private void SaveHtmlTopic(Section section, Topic topic)
        {
            Document dummyDoc = new Document();
            dummyDoc.RemoveAllChildren();
            dummyDoc.AppendChild(dummyDoc.ImportNode(section, true, ImportFormatMode.KeepSourceFormatting));

            dummyDoc.BuiltInDocumentProperties.Title = topic.Title;

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                PrettyFormat = true,
                AllowNegativeIndent = true, // This is to allow headings to appear to the left of the main text.
                ExportHeadersFootersMode = ExportHeadersFootersMode.None
            };

            dummyDoc.Save(topic.FileName, saveOptions);
        }

        /// <summary>
        /// Generates a table of contents for the topics and saves to contents .html.
        /// </summary>
        private void SaveTableOfContents(List<Topic> topics)
        {
            Document tocDoc = new Document(mTocTemplate);

            // We use a custom mail merge event handler defined below,
            // and a custom mail merge data source based on collecting the topics we created.
            tocDoc.MailMerge.FieldMergingCallback = new HandleTocMergeField();
            tocDoc.MailMerge.ExecuteWithRegions(new TocMailMergeDataSource(topics));

            tocDoc.Save(Path.Combine(mDstDir, "contents.html"));
        }

        private class HandleTocMergeField : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (mBuilder == null)
                    mBuilder = new DocumentBuilder(e.Document);

                // Our custom data source returns topic objects.
                Topic topic = (Topic) e.FieldValue;

                mBuilder.MoveToMergeField(e.FieldName);
                mBuilder.InsertHyperlink(topic.Title, topic.FileName, false);

                // Signal to the mail merge engine that it does not need to insert text into the field.
                e.Text = "";
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }

            private DocumentBuilder mBuilder;
        }

        private Document mDoc;
        private string mTocTemplate;
        private string mDstDir;
    }

    internal class Topic
    {
        internal Topic(string title, string fileName)
        {
            Title = title;
            FileName = fileName;
        }

        internal string Title { get; }

        internal string FileName { get; }
    }

    internal class TocMailMergeDataSource : IMailMergeDataSource
    {
        internal TocMailMergeDataSource(List<Topic> topics)
        {
            mTopics = topics;
            mIndex = -1;
        }

        public bool MoveNext()
        {
            if (mIndex < mTopics.Count - 1)
            {
                mIndex++;
                return true;
            }

            return false;
        }

        public bool GetValue(string fieldName, out object fieldValue)
        {
            if (fieldName == "TocEntry")
            {
                // The template document is supposed to have only one field called "TocEntry".
                fieldValue = mTopics[mIndex];
                return true;
            }

            fieldValue = null;
            return false;
        }

        public string TableName => "TOC";

        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            return null;
        }

        private readonly List<Topic> mTopics;
        private int mIndex;
    }
}