//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////


using NUnit.Framework;
using System;
using System.Collections;

using Aspose.Words;
using Aspose.Words.Fields;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading;

namespace Examples
{
    [TestFixture]
    public class ExField : ExBase
    {
        [Test]
        public void UpdateTOC()
        {
            Document doc = new Document();

            //ExStart
            //ExId:UpdateTOC
            //ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
            doc.UpdateFields();
            //ExEnd
        }

        [Test]
        public void GetFieldType()
        {
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            //ExStart
            //ExFor:FieldType
            //ExFor:FieldChar
            //ExFor:FieldChar.FieldType
            //ExSummary:Shows how to find the type of field that is represented by a node which is derived from FieldChar.
            FieldChar fieldStart = (FieldChar)doc.GetChild(NodeType.FieldStart, 0, true);
            FieldType type = fieldStart.FieldType;
            //ExEnd
        }

        [Test]
        public void GetFieldFromDocument()
        {
            //ExStart
            //ExFor:FieldChar.GetField
            //ExId:GetField
            //ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            FieldStart fieldStart = (FieldStart)doc.GetChild(NodeType.FieldStart, 0, true);

            // Retrieve the facade object which represents the field in the document.
            Field field = fieldStart.GetField();
            
            Console.WriteLine("Field code:" + field.GetFieldCode());
            Console.WriteLine("Field result: " + field.Result);
            Console.WriteLine("Is locked: " + field.IsLocked);

            // This updates only this field in the document.
            field.Update();
            //ExEnd
        }

        [Test]
        public void GetFieldFromFieldCollection()
        {
            //ExStart
            //ExId:GetFieldFromFieldCollection
            //ExSummary:Demonstrates how to retrieve a field using the range of a node.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            Field field = doc.Range.Fields[0];

            // This should be the first field in the document - a TOC field.
            Console.WriteLine(field.Type);
            //ExEnd
        }

        [Test]
        public void InsertTCField()
        {
            //ExStart
            //ExId:InsertTCField
            //ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
            // Create a blank document.
            Document doc = new Document();

            // Create a document builder to insert content with.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TC field at the current document builder position.
            builder.InsertField("TC \"Entry Text\" \\f t");
            //ExEnd
        }

        [Test]
        public void ChangeLocale()
        {
            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder b = new DocumentBuilder(doc);
            b.InsertField("MERGEFIELD Date");

            //ExStart
            //ExId:ChangeCurrentCulture
            //ExSummary:Shows how to change the culture used in formatting fields during update.
            // Store the current culture so it can be set back once mail merge is complete.
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            // Set to German language so dates and numbers are formatted using this culture during mail merge.
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            // Execute mail merge.
            doc.MailMerge.Execute(new string[] { "Date" }, new object[] { DateTime.Now });

            // Restore the original culture.
            Thread.CurrentThread.CurrentCulture = currentCulture;
            //ExEnd

            doc.Save(MyDir + "Field.ChangeLocale Out.doc");
        }

        [Test]
        public void RemoveTOCFromDocument()
        {
            //ExStart
            //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
            //ExId:RemoveTableOfContents
            //ExSummary:Demonstrates how to remove a specified TOC from a document.
            // Open a document which contains a TOC.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            // Remove the first TOC from the document.
            Field tocField = doc.Range.Fields[0];
            tocField.Remove();

            // Save the output.
            doc.Save(MyDir + "Document.TableOfContentsRemoveTOC Out.doc");
            //ExEnd
        }

        [Test]
        //ExStart
        //ExId:TCFieldsRangeReplace
        //ExSummary:Shows how to find and insert a TC field at text in a document. 
        public void InsertTCFieldsAtText()
        {
            Document doc = new Document();

            // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
            doc.Range.Replace(new Regex("The Beginning"), new InsertTCFieldHandler("Chapter 1", "\\l 1"), false);
        }

        public class InsertTCFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields.
            private string mFieldText;
            private string mFieldSwitches;

            /// <summary>
            /// The switches to use for each TC field. Can be an empty string or null.
            /// </summary>
            public InsertTCFieldHandler(string switches) : this(string.Empty, switches)
            {
                mFieldSwitches = switches;
            }

            /// <summary>
            /// The display text and switches to use for each TC field. Display name can be an empty string or null.
            /// </summary>
            public InsertTCFieldHandler(string text, string switches)
            {
                mFieldText = text;
                mFieldSwitches = switches;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                // Create a builder to insert the field.
                DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
                // Move to the first node of the match.
                builder.MoveTo(args.MatchNode);

                // If the user specified text to be used in the field as display text then use that, otherwise use the 
                // match string as the display text.
                string insertText;

                if (!string.IsNullOrEmpty(mFieldText))
                    insertText = mFieldText;
                else
                    insertText = args.Match.Value;

                // Insert the TC field before this node using the specified string as the display text and user defined switches.
                builder.InsertField(string.Format("TC \"{0}\" {1}", insertText, mFieldSwitches));

                // We have done what we want so skip replacement.
                return ReplaceAction.Skip;
            }
        }
        //ExEnd
    }
}
