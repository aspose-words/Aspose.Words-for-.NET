using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using System.Text.RegularExpressions;
using System.Text;
namespace CSharp.Programming_Documents.Working_with_Hyperlink
{
    class ReplaceHyperlinks
    {        
        public static void Run()
        {
            //ExStart:ReplaceHyperlinks
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithHyperlink();
            string NewUrl = @"http://www.aspose.com";
            string NewName = "Aspose - The .NET & Java Component Publisher";
            Document doc = new Document(dataDir + "ReplaceHyperlinks.doc");

            // Hyperlinks in a Word documents are fields, select all field start nodes so we can find the hyperlinks.
            NodeList fieldStarts = doc.SelectNodes("//FieldStart");
            foreach (FieldStart fieldStart in fieldStarts)
            {
                if (fieldStart.FieldType.Equals(FieldType.FieldHyperlink))
                {
                    // The field is a hyperlink field, use the "facade" class to help to deal with the field.
                    Hyperlink hyperlink = new Hyperlink(fieldStart);

                    // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                    if (hyperlink.IsLocal)
                        continue;

                    // The Hyperlink class allows to set the target URL and the display name
                    // of the link easily by setting the properties.
                    hyperlink.Target = NewUrl;
                    hyperlink.Name = NewName;
                }
            }
            dataDir = dataDir + "ReplaceHyperlinks_out_.doc";
            doc.Save(dataDir);
            //ExEnd:ReplaceHyperlinks
            Console.WriteLine("\nHyperlinks replaced successfully.\nFile saved at " + dataDir);
        }
        //ExStart:Hyperlink
        /// <summary>
        /// This "facade" class makes it easier to work with a hyperlink field in a Word document.
        ///
        /// A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words
        /// consists of several nodes and it might be difficult to work with all those nodes directly.
        /// Note this is a simple implementation and will work only if the hyperlink code and name
        /// each consist of one Run only.
        ///
        /// [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
        ///
        /// The field code contains a string in one of these formats:
        /// HYPERLINK "url"
        /// HYPERLINK \l "bookmark name"
        ///
        /// The field result contains text that is displayed to the user.
        /// </summary>
        internal class Hyperlink
        {
            internal Hyperlink(FieldStart fieldStart)
            {
                if (fieldStart == null)
                    throw new ArgumentNullException("fieldStart");
                if (!fieldStart.FieldType.Equals(FieldType.FieldHyperlink))
                    throw new ArgumentException("Field start type must be FieldHyperlink.");

                mFieldStart = fieldStart;

                // Find the field separator node.
                mFieldSeparator = fieldStart.GetField().Separator;
                if (mFieldSeparator == null)
                    throw new InvalidOperationException("Cannot find field separator.");

                mFieldEnd = fieldStart.GetField().End;

                // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
                string fieldCode = fieldStart.GetField().GetFieldCode();
                Match match = gRegex.Match(fieldCode.Trim());
                mIsLocal = (match.Groups[1].Length > 0);    //The link is local if \l is present in the field code.
                mTarget = match.Groups[2].Value;
            }

            /// <summary>
            /// Gets or sets the display name of the hyperlink.
            /// </summary>
            internal string Name
            {
                get
                {
                    return GetTextSameParent(mFieldSeparator, mFieldEnd);
                }
                set
                {
                    // Hyperlink display name is stored in the field result which is a Run
                    // node between field separator and field end.
                    Run fieldResult = (Run)mFieldSeparator.NextSibling;
                    fieldResult.Text = value;

                    // But sometimes the field result can consist of more than one run, delete these runs.
                    RemoveSameParent(fieldResult.NextSibling, mFieldEnd);
                }
            }

            /// <summary>
            /// Gets or sets the target url or bookmark name of the hyperlink.
            /// </summary>
            internal string Target
            {
                get
                {
                    string dummy = null;  // This is needed to fool the C# to VB.NET converter.
                    return mTarget;
                }
                set
                {
                    mTarget = value;
                    UpdateFieldCode();
                }
            }

            /// <summary>
            /// True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
            /// </summary>
            internal bool IsLocal
            {
                get
                {
                    return mIsLocal;
                }
                set
                {
                    mIsLocal = value;
                    UpdateFieldCode();
                }
            }

            private void UpdateFieldCode()
            {
                // Field code is stored in a Run node between field start and field separator.
                Run fieldCode = (Run)mFieldStart.NextSibling;
                fieldCode.Text = string.Format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget);

                // But sometimes the field code can consist of more than one run, delete these runs.
                RemoveSameParent(fieldCode.NextSibling, mFieldSeparator);
            }

            /// <summary>
            /// Retrieves text from start up to but not including the end node.
            /// </summary>
            private static string GetTextSameParent(Node startNode, Node endNode)
            {
                if ((endNode != null) && (startNode.ParentNode != endNode.ParentNode))
                    throw new ArgumentException("Start and end nodes are expected to have the same parent.");

                StringBuilder builder = new StringBuilder();
                for (Node child = startNode; !child.Equals(endNode); child = child.NextSibling)
                    builder.Append(child.GetText());

                return builder.ToString();
            }

            /// <summary>
            /// Removes nodes from start up to but not including the end node.
            /// Start and end are assumed to have the same parent.
            /// </summary>
            private static void RemoveSameParent(Node startNode, Node endNode)
            {
                if (((endNode != null) && (startNode != null)) && (startNode.ParentNode != endNode.ParentNode))
                    throw new ArgumentException("Start and end nodes are expected to have the same parent.");

                Node curChild = startNode;
                while ((curChild != null) && (curChild != endNode))
                {
                    Node nextChild = curChild.NextSibling;
                    curChild.Remove();
                    curChild = nextChild;
                }
            }

            private readonly Node mFieldStart;
            private readonly Node mFieldSeparator;
            private readonly Node mFieldEnd;
            private bool mIsLocal;
            private string mTarget;

            /// <summary>
            /// RK I am notoriously bad at regexes. It seems I don't understand their way of thinking.
            /// </summary>
            private static readonly Regex gRegex = new Regex(
                "\\S+" +            // one or more non spaces HYPERLINK or other word in other languages
                "\\s+" +            // one or more spaces
                "(?:\"\"\\s+)?" +    // non capturing optional "" and one or more spaces, found in one of the customers files.
                "(\\\\l\\s+)?" +    // optional \l flag followed by one or more spaces
                "\"" +                // one apostrophe
                "([^\"]+)" +        // one or more chars except apostrophe (hyperlink target)
                "\""                // one closing apostrophe
                );
        }
        //ExEnd:Hyperlink
    }
}
