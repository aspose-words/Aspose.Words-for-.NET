// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Properties;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExProperties : ApiExampleBase
    {
        [Test]
        public void EnumerateProperties()
        {
            //ExStart
            //ExFor:Document.BuiltInDocumentProperties
            //ExFor:Document.CustomDocumentProperties
            //ExFor:BuiltInDocumentProperties
            //ExFor:CustomDocumentProperties
            //ExId:DocumentProperties
            //ExSummary:Enumerates through all built-in and custom properties in a document.
            Document doc = new Document(MyDir + "Properties.doc");

            Console.WriteLine("1. Document name: {0}", doc.OriginalFileName);

            Console.WriteLine("2. Built-in Properties");
            foreach (DocumentProperty docProperty in doc.BuiltInDocumentProperties)
                Console.WriteLine("{0} : {1}", docProperty.Name, docProperty.Value);

            Console.WriteLine("3. Custom Properties");
            foreach (DocumentProperty docProperty in doc.CustomDocumentProperties)
                Console.WriteLine("{0} : {1}", docProperty.Name, docProperty.Value);
            //ExEnd
        }

        [Test]
        public void EnumeratePropertiesWithIndexer()
        {
            //ExStart
            //ExFor:DocumentPropertyCollection.Count
            //ExFor:DocumentPropertyCollection.Item(int)
            //ExFor:DocumentProperty
            //ExFor:DocumentProperty.Name
            //ExFor:DocumentProperty.Value
            //ExFor:DocumentProperty.Type
            //ExSummary:Enumerates through all built-in and custom properties in a document using indexed access.
            Document doc = new Document(MyDir + "Properties.doc");

            Console.WriteLine("1. Document name: {0}", doc.OriginalFileName);

            Console.WriteLine("2. Built-in Properties");
            for (int i = 0; i < doc.BuiltInDocumentProperties.Count; i++)
            {
                DocumentProperty docProperty = doc.BuiltInDocumentProperties[i];
                Console.WriteLine("{0}({1}) : {2}", docProperty.Name, docProperty.Type, docProperty.Value);
            }

            Console.WriteLine("3. Custom Properties");
            for (int i = 0; i < doc.CustomDocumentProperties.Count; i++)
            {
                DocumentProperty docProperty = doc.CustomDocumentProperties[i];
                Console.WriteLine("{0}({1}) : {2}", docProperty.Name, docProperty.Type, docProperty.Value);
            }
            //ExEnd
        }

        [Test]
        public void BuiltInNamedAccess()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.Item(String)
            //ExFor:DocumentProperty.ToString
            //ExSummary:Retrieves a built-in document property by name.
            Document doc = new Document(MyDir + "Properties.doc");

            DocumentProperty docProperty = doc.BuiltInDocumentProperties["Keywords"];
            Console.WriteLine(docProperty.ToString());
            //ExEnd
        }

        [Test]
        public void BuiltInPropertiesDirectAccess()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.Author
            //ExFor:BuiltInDocumentProperties.Category
            //ExFor:BuiltInDocumentProperties.Comments
            //ExFor:BuiltInDocumentProperties.Company
            //ExFor:BuiltInDocumentProperties.CreatedTime
            //ExFor:BuiltInDocumentProperties.Keywords
            //ExFor:BuiltInDocumentProperties.LastPrinted
            //ExFor:BuiltInDocumentProperties.LastSavedBy
            //ExFor:BuiltInDocumentProperties.LastSavedTime
            //ExFor:BuiltInDocumentProperties.Manager
            //ExFor:BuiltInDocumentProperties.NameOfApplication
            //ExFor:BuiltInDocumentProperties.RevisionNumber
            //ExFor:BuiltInDocumentProperties.Subject
            //ExFor:BuiltInDocumentProperties.Template
            //ExFor:BuiltInDocumentProperties.Title
            //ExFor:BuiltInDocumentProperties.TotalEditingTime
            //ExFor:BuiltInDocumentProperties.Version
            //ExSummary:Retrieves information from the built-in document properties.
            String fileName = MyDir + "Properties.doc";
            Document doc = new Document(fileName);

            Console.WriteLine("Document name: {0}", fileName);
            Console.WriteLine("Document author: {0}", doc.BuiltInDocumentProperties.Author);
            Console.WriteLine("Category: {0}", doc.BuiltInDocumentProperties.Category);
            Console.WriteLine("Comments: {0}", doc.BuiltInDocumentProperties.Comments);
            Console.WriteLine("Company: {0}", doc.BuiltInDocumentProperties.Company);
            Console.WriteLine("Create time: {0}", doc.BuiltInDocumentProperties.CreatedTime);
            Console.WriteLine("Keywords: {0}", doc.BuiltInDocumentProperties.Keywords);
            Console.WriteLine("Last printed: {0}", doc.BuiltInDocumentProperties.LastPrinted);
            Console.WriteLine("Last saved by: {0}", doc.BuiltInDocumentProperties.LastSavedBy);
            Console.WriteLine("Last saved: {0}", doc.BuiltInDocumentProperties.LastSavedTime);
            Console.WriteLine("Lines: {0}", doc.BuiltInDocumentProperties.Lines);
            Console.WriteLine("Manager: {0}", doc.BuiltInDocumentProperties.Manager);
            Console.WriteLine("Name of application: {0}", doc.BuiltInDocumentProperties.NameOfApplication);
            Console.WriteLine("Revision number: {0}", doc.BuiltInDocumentProperties.RevisionNumber);
            Console.WriteLine("Subject: {0}", doc.BuiltInDocumentProperties.Subject);
            Console.WriteLine("Template: {0}", doc.BuiltInDocumentProperties.Template);
            Console.WriteLine("Title: {0}", doc.BuiltInDocumentProperties.Title);
            Console.WriteLine("Total editing time: {0}", doc.BuiltInDocumentProperties.TotalEditingTime);
            Console.WriteLine("Version: {0}", doc.BuiltInDocumentProperties.Version);
            //ExEnd
        }

        //ExStart
        //ExFor:BuiltInDocumentProperties.Bytes
        //ExFor:BuiltInDocumentProperties.Characters
        //ExFor:BuiltInDocumentProperties.CharactersWithSpaces
        //ExFor:BuiltInDocumentProperties.ContentStatus
        //ExFor:BuiltInDocumentProperties.ContentType
        //ExFor:BuiltInDocumentProperties.Lines
        //ExFor:BuiltInDocumentProperties.Pages
        //ExFor:BuiltInDocumentProperties.Paragraphs
        //ExFor:BuiltInDocumentProperties.Words
        //ExSummary:Shows how to work with document properties from the "Content" category.
        [Test] //ExSkip
        public void BuiltInPropertiesContent()
        {
            // Create a new document and populate it with paragraphs and text
            Document doc = new Document();
            doc.RemoveAllChildren();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Paragraph 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.Write("Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
            builder.Writeln("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Paragraph 3.");

            // A document's built in properties can be viewed without opening the document,
            // by right-clicking the file in Windows Explorer, navigating to Properties > Details
            // The "Content" category will have the properties we are about to work with
            // They are not updated automatically; their values need to be sourced from the document and assigned manually
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            // Page count: The number of pages is in a document's PageCount attribute, which is always up to date,
            // and can be transferred to a built in property at any time
            properties.Pages = doc.PageCount;
            Assert.AreEqual(3, properties.Pages);

            // Word count: The UpdateWordCount() automatically assigns the real time word/character counts to the relevant properties
            doc.UpdateWordCount();
            Assert.AreEqual(54, properties.Words);
            Assert.AreEqual(305, properties.Characters);
            Assert.AreEqual(356, properties.CharactersWithSpaces);

            // Line count: Count the lines in a document and assign value to the property
            LineCounter lineCounter = new LineCounter(doc);
            properties.Lines = lineCounter.GetLineCount();
            Assert.AreEqual(6, properties.Lines);

            // Paragraph count: Assign the size of the array of all paragraphs to the Paragraphs property
            properties.Paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Count;
            Assert.AreEqual(4, properties.Paragraphs);

            // Bytes: Get the real document file size from a stream and assign it to the property
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Docx);
                properties.Bytes = (int)stream.Length;
                Assert.AreEqual(5756, properties.Bytes);
            }
            
            // Template: If we change a document's template, the Template property will need a descriptive name for the template added manually
            doc.AttachedTemplate = MyDir + "Document.BusinessBrochureTemplate.dot";
            Assert.AreEqual("Normal.dot", properties.Template);          
            properties.Template = "Document.BusinessBrochureTemplate.dot";

            // Content status: This is a descriptive field that we can glance at without opening the document
            properties.ContentStatus = "Draft";

            // Content type: Upon saving, any value we assign to this field will be overwritten by the MIME type of the output save format
            Assert.AreEqual("", properties.ContentType);

            // If the document contains links that are up to date, we can set this to true
            Assert.False(properties.LinksUpToDate);
            
            doc.Save(ArtifactsDir + "Properties.BuiltInPropertiesContent.docx");
        }

        /// <summary>
        /// Util class that counts the lines in a document
        /// Upon construction, traverses the document's layout entities tree, counting entities of the "Line" type that also contain real text
        /// </summary>
        private class LineCounter
        {
            public LineCounter(Document doc)
            {
                mLayoutEnumerator = new LayoutEnumerator(doc);

                CountLines();
            }

            public int GetLineCount()
            {
                return mLineCount;
            }

            private void CountLines()
            {
                do
                {
                    if (mLayoutEnumerator.Type == LayoutEntityType.Line)
                    {
                        mScanningLineForRealText = true;
                    }

                    if (mLayoutEnumerator.MoveFirstChild())
                    {
                        if (mScanningLineForRealText && mLayoutEnumerator.Kind.StartsWith("TEXT"))
                        {
                            mLineCount++;
                            mScanningLineForRealText = false;
                        }
                        CountLines();
                        mLayoutEnumerator.MoveParent();
                    }
                } while (mLayoutEnumerator.MoveNext());
            }

            private readonly LayoutEnumerator mLayoutEnumerator;
            private int mLineCount;
            private bool mScanningLineForRealText;
        }
        //ExEnd

        [Test]
        public void CustomNamedAccess()
        {
            //ExStart
            //ExFor:DocumentPropertyCollection.Item(String)
            //ExFor:CustomDocumentProperties.Add(String,DateTime)
            //ExFor:DocumentProperty.ToDateTime
            //ExSummary:Retrieves a custom document property by name.
            Document doc = new Document(MyDir + "Properties.doc");

            DocumentProperty docProperty = doc.CustomDocumentProperties["Authorized Date"];

            if (docProperty != null)
            {
                Console.WriteLine(docProperty.ToDateTime());
            }
            else
            {
                Console.WriteLine("The document is not authorized. Authorizing...");
                doc.CustomDocumentProperties.Add("AuthorizedDate", DateTime.Now);
            }

            //ExEnd
        }

        [Test]
        public void CustomAdd()
        {
            //ExStart
            //ExFor:CustomDocumentProperties.Add(String,String)
            //ExFor:CustomDocumentProperties.Add(String,Boolean)
            //ExFor:CustomDocumentProperties.Add(String,int)
            //ExFor:CustomDocumentProperties.Add(String,DateTime)
            //ExFor:CustomDocumentProperties.Add(String,Double)
            //ExId:AddCustomProperties
            //ExSummary:Checks if a custom property with a given name exists in a document and adds few more custom document properties.
            Document doc = new Document(MyDir + "Properties.doc");

            CustomDocumentProperties docProperties = doc.CustomDocumentProperties;

            if (docProperties["Authorized"] == null)
            {
                docProperties.Add("Authorized", true);
                docProperties.Add("Authorized By", "John Smith");
                docProperties.Add("Authorized Date", DateTime.Today);
                docProperties.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
                docProperties.Add("Authorized Amount", 123.45);
            }

            //ExEnd
        }

        [Test]
        public void CustomRemove()
        {
            //ExStart
            //ExFor:DocumentPropertyCollection.Remove
            //ExId:RemoveCustomProperties
            //ExSummary:Removes a custom document property.
            Document doc = new Document(MyDir + "Properties.doc");

            doc.CustomDocumentProperties.Remove("Authorized Date");
            //ExEnd
        }

        [Test]
        public void PropertyTypes()
        {
            //ExStart
            //ExFor:DocumentProperty.Type
            //ExFor:DocumentProperty.ToBool
            //ExFor:DocumentProperty.ToInt
            //ExFor:DocumentProperty.ToDouble
            //ExFor:DocumentProperty.ToString
            //ExFor:DocumentProperty.ToDateTime
            //ExFor:PropertyType
            //ExSummary:Retrieves the types and values of the custom document properties.
            Document doc = new Document(MyDir + "Properties.doc");

            foreach (DocumentProperty docProperty in doc.CustomDocumentProperties)
            {
                Console.WriteLine(docProperty.Name);
                switch (docProperty.Type)
                {
                    case PropertyType.String:
                        Console.WriteLine("It's a String value.");
                        Console.WriteLine(docProperty.ToString());
                        break;
                    case PropertyType.Boolean:
                        Console.WriteLine("It's a boolean value.");
                        Console.WriteLine(docProperty.ToBool());
                        break;
                    case PropertyType.Number:
                        Console.WriteLine("It's an integer value.");
                        Console.WriteLine(docProperty.ToInt());
                        break;
                    case PropertyType.DateTime:
                        Console.WriteLine("It's a date time value.");
                        Console.WriteLine(docProperty.ToDateTime());
                        break;
                    case PropertyType.Double:
                        Console.WriteLine("It's a double value.");
                        Console.WriteLine(docProperty.ToDouble());
                        break;
                    case PropertyType.Other:
                        Console.WriteLine("Other value.");
                        break;
                    default:
                        throw new Exception("Unknown property type.");
                }
            }

            //ExEnd
        }
    }
}