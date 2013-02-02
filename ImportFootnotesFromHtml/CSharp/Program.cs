//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// 3/1/08 by Roman Korchagin
using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;


namespace ImportFootnotesFromHtml
{
    /// <summary>
    /// This is a sample code for http://www.aspose.com/Community/Forums/thread/107477.aspx
    /// 
    /// The scenario is as follows:
    /// 
    /// 1. The customer has a DOC file with footnotes.
    /// 
    /// 2. The customer uses Aspose.Words to convert DOC to HTML. Aspose.Words converts
    /// footnotes and endnotes into hyperlinks. There are two hyperlinks per footnote actually.
    /// One link is "forward" from the main text to the text of the footnote. 
    /// Another is "backward" from the text of the footnote to the main text.
    /// 
    /// 3. The customer uses Aspose.Words to convert HTML back to DOC.
    /// In the current version of Aspose.Words the hyperlinks do not become footnotes,
    /// they just stay as hyperlink fields in the document. The customer wants 
    /// original footnotes to become footnotes during DOC->HTML->DOC roundtrip.
    /// 
    /// This code is a workaround that detects hyperlinks related to footnotes and converts
    /// them into proper footnotes. At some point in the future, this code will not be needed
    /// when Aspose.Words will guarantee footnotes roundtripping.
    /// 
    /// This code demonstrates some useful techniques, such as enumerating over nodes,
    /// getting field code, removing fields etc.
    /// </summary>
    class Program
    {
        public static void Main(string[] args)
        {
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Load DOC with footnotes into a document object.
            Document srcDoc = new Document(Path.Combine(dataDir, "FootnoteSample.doc"));

            // Save to HTML file. Footnotes get converted to hyperlinks.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.PrettyFormat = true;
            string htmlFile = Path.Combine(dataDir, "FootnoteSample Out.html");
            srcDoc.Save(htmlFile, saveOptions);

            // Load HTML back into a document object. 
            // In the current version of Aspose.Words hyperlinks do not become footnotes again,
            // they become regular hyperlinks.
            Document dstDoc = new Document(htmlFile);

            // You can open this document in MS Word and see there are no footnotes, just hyperlinks.
            dstDoc.Save(Path.Combine(dataDir, "FootnoteSample Out1.doc"));

            // This is the workaround method I'm suggesting. It will recognize hyperlinks that
            // should become footnotes and convert them into footnotes.
            ConvertHyperlinksToFootnotes(dstDoc);

            // You can open this document in MS Word and see the footnotes are as expected.
            dstDoc.Save(Path.Combine(dataDir, "FootnoteSample Out2.doc"));
        }

        /// <summary>
        /// A "workaround" method that you can use after DOC->HTML->DOC conversion of a document
        /// with footnotes. Will make sure that original DOC footnotes will still be footnotes in 
        /// the final DOC file.
        /// </summary>
        internal static void ConvertHyperlinksToFootnotes(Document doc)
        {
            // When processing HYPERLINK fields we will remove them (convert to footnotes).
            // Since it is not a good thing to delete nodes while iterating over a collection, 
            // we will collect the nodes during the first pass and delete them during the second.
            //
            // These collections contain HYPERLINK field starts of footnotes and endnotes in the main document.
            Hashtable ftnFieldStarts = new Hashtable();
            Hashtable ednFieldStarts = new Hashtable();
            // These collections contain HYPERLINK field starts of footnotes and endnotes themselves.
            Hashtable ftnRefFieldStarts = new Hashtable();
            Hashtable ednRefFieldStarts = new Hashtable();

            // Collect all the nodes into arrays before we start deleting them.
            CollectFieldStarts(doc, ftnFieldStarts, ednFieldStarts, ftnRefFieldStarts, ednRefFieldStarts);

            // Remove the HR shapes that separate footnotes and endnotes from the main text.
            RemoveHorizontalLine(ftnRefFieldStarts);
            RemoveHorizontalLine(ednRefFieldStarts);

            // Convert the HYPERLINK fields into proper footnotes and endnotes.
            ConvertFieldsToNotes(ftnFieldStarts, ftnRefFieldStarts, FootnoteType.Footnote);
            ConvertFieldsToNotes(ednFieldStarts, ednRefFieldStarts, FootnoteType.Endnote);
        }

        /// <summary>
        /// Collects field start nodes of HYPERLINK fields related to footnotes and endnotes.
        /// </summary>
        /// <param name="doc">The document to process.</param>
        /// <param name="ftnFieldStarts">Starts of HYPERLINK fields that represent footnotes will be returned here.</param>
        /// <param name="ednFieldStarts">Start of HYPERLINK fields that represent endnotes will be returned here.</param>
        /// <param name="ftnRefFieldStarts">Starts of HYPERLINK fields that are back-links to footnotes will be returned here.</param>
        /// <param name="ednRefFieldStarts">Starts of HYPERLINK fields that are back-links to endnotes will be returned here.</param>
        private static void CollectFieldStarts(
            Document doc,
            Hashtable ftnFieldStarts,
            Hashtable ednFieldStarts,
            Hashtable ftnRefFieldStarts,
            Hashtable ednRefFieldStarts)
        {
            // This regex parses the "command" which we use to determine the footnote/endnote type
            // and the id.
            Regex regex = new Regex(@"HYPERLINK \\l \""(?<cmd>(_ftn|_edn|_ftnref|_ednref))(?<id>[0-9]+)\""");

            // We need to process all HYPERLINK fields. Therefore select all field starts.
            NodeCollection fieldStarts = doc.GetChildNodes(NodeType.FieldStart, true);
            foreach (FieldStart fieldStart in fieldStarts)
            {
                if (fieldStart.FieldType == FieldType.FieldHyperlink)
                {
                    // The field is a hyperlink, lets analyze the field code.
                    string fieldCode = GetFieldCode(fieldStart);

                    Match match = regex.Match(fieldCode);
                    string cmd = match.Groups["cmd"].Value;
                    string id = match.Groups["id"].Value;

                    switch (cmd)
                    {
                        case "_ftn":
                            // Field is HYPERLINK \l "_ftn1". It is a footnote in the main document.
                            ftnFieldStarts.Add(int.Parse(id), fieldStart);
                            break;
                        case "_edn":
                            ednFieldStarts.Add(int.Parse(id), fieldStart);
                            break;
                        case "_ftnref":
                            // Field is HYPERLINK \l "_ftnref1". It is a back-link to the footnote in 
                            // the main document. The parent paragraph contains the text of the footnote.
                            ftnRefFieldStarts.Add(int.Parse(id), fieldStart);
                            break;
                        case "_ednref":
                            ednRefFieldStarts.Add(int.Parse(id), fieldStart);
                            break;
                        default:
                            // Do nothing.
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// A simplistic method to get the field code as a string.
        /// Goes trough all Run nodes after the field start and concatenates their text.
        /// </summary>
        private static string GetFieldCode(FieldStart fieldStart)
        {
            StringBuilder fieldCode = new StringBuilder();
            Node curNode = fieldStart.NextSibling;
            while (curNode is Run)
            {
                fieldCode.Append(curNode.GetText());
                curNode = curNode.NextSibling;                
            }
            return fieldCode.ToString();
        }

        /// <summary>
        /// Performs the actual conversion of HYPERLINK fields into footnotes/endnote.
        /// </summary>
        /// <param name="noteFieldStarts">The starts of hyperlink fields in the main document.</param>
        /// <param name="refNoteFieldStarts">The starts of back-link hyperlink fields in the footnotes.</param>
        /// <param name="noteType">Specifies whether we are processing footnotes or endnotes.</param>
        private static void ConvertFieldsToNotes(
            Hashtable noteFieldStarts, 
            Hashtable refNoteFieldStarts, 
            FootnoteType noteType)
        {
            foreach (DictionaryEntry entry in noteFieldStarts)
            {
                // Footnote/endnote id is stored in the key.
                int id = (int)entry.Key;
                FieldStart noteFieldStart = (FieldStart)entry.Value;
                // Using the id we can retrieve the field start of the back-link field.
                FieldStart refNoteFieldStart = (FieldStart)refNoteFieldStarts[id];

                ConvertFieldToNote(noteFieldStart, refNoteFieldStart, noteType);
            }   
        }

        /// <summary>
        /// Performs the actual task of converting one HYPERLINK into footnote or endnote.
        /// </summary>
        /// <param name="noteFieldStart">The start of the hyperlink field in the main document.</param>
        /// <param name="refNoteFieldStart">The start of the back-link hyperlink field in the footnote.</param>
        /// <param name="noteType">Specifies whether we are processing footnotes or endnotes.</param>
        private static void ConvertFieldToNote(
            FieldStart noteFieldStart, 
            FieldStart refNoteFieldStart, 
            FootnoteType noteType)
        {
            // This is the paragraph that contains the text of the footnote.
            Paragraph oldNotePara = refNoteFieldStart.ParentParagraph;
            
            // Delete the hyperlink field from the text of the footnote because we don't need it anymore.
            DeleteField(refNoteFieldStart);

            // Use document builder to move to the place in the main document where the footnote
            // should be and insert a proper footnote.
            DocumentBuilder builder = new DocumentBuilder((Document)noteFieldStart.Document);
            builder.MoveTo(noteFieldStart);
            Footnote note = builder.InsertFootnote(noteType, "");

            // Move all content from the old footnote paragraphs into the new.
            Paragraph newNotePara = note.FirstParagraph;
            Node curNode = oldNotePara.FirstChild;
            while (curNode != null)
            {
                Node nextNode = curNode.NextSibling;
                newNotePara.AppendChild(curNode);
                curNode = nextNode;
            }

            // Delete the old paragraph that represented the footnote. 
            oldNotePara.Remove();

            // Remove the hyperlink field from the main text to the footnote.
            DeleteField(noteFieldStart);
        }

        /// <summary>
        /// A simplistic method to delete all nodes of a field given a field start node.
        /// </summary>
        private static void DeleteField(FieldStart fieldStart)
        {
            Node curNode = fieldStart;
            while (curNode.NodeType != NodeType.FieldEnd)
            {
                Node nextNode = curNode.NextSibling;
                curNode.Remove();
                curNode = nextNode;
            }
            curNode.Remove();
        }

        /// <summary>
        /// There is an HR (horizontal rule) shape in a separate paragraph just before
        /// the first footnote and first endnote in a document imported from HTML.
        /// This method deletes the paragraph and the HR shape.
        /// </summary>
        private static void RemoveHorizontalLine(Hashtable noteRefFieldStarts)
        {
            // Footnote and endnote ids start from 1. Therefore we can get the first note.
            FieldStart noteFieldStart = (FieldStart)noteRefFieldStarts[1];
            // This is the paragraph that contains the first footnote.
            Paragraph notePara = noteFieldStart.ParentParagraph;
            // This is the previous paragraph that contains the HR shape. Delete the paragraph.
            Paragraph hrPara = (Paragraph)notePara.PreviousSibling;
            hrPara.Remove();
        }
    }
}
