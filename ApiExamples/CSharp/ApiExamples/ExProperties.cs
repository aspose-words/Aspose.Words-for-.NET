// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Web.UI;
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
        //ExFor:BuiltInDocumentProperties.LinksUpToDate
        //ExFor:BuiltInDocumentProperties.Pages
        //ExFor:BuiltInDocumentProperties.Paragraphs
        //ExFor:BuiltInDocumentProperties.Words
        //ExSummary:Shows how to work with document properties from the "Content" category.
        [Test] //ExSkip
        public void Content()
        {
            // Open a document with a couple paragraphs of content
            Document doc = new Document(MyDir + "Properties.Content.docx");

            // If we want to display document stats such as page/word counts inside a document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc
            // Also, these statistics are found in File > Properties > Advanced Properties > Statistics
            // To be able to glance at these values without opening the document, we can use a document's built in document property collection
            // These properties are accessed by right-clicking the file in Windows Explorer and navigating to Properties > Details
            // The "Content" category will have all the properties we will work with
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            // Page count: The PageCount attribute shows the page count in real time and its value can be assigned to the Pages property
            properties.Pages = doc.PageCount;
            Assert.AreEqual(2, properties.Pages);

            // Word count: The UpdateWordCount() automatically assigns the real time word/character counts to the respective built in properties
            doc.UpdateWordCount();
            Assert.AreEqual(198, properties.Words);
            Assert.AreEqual(1114, properties.Characters);
            Assert.AreEqual(1310, properties.CharactersWithSpaces);

            // Line count: Count the lines in a document and assign value to the Lines property
            LineCounter lineCounter = new LineCounter(doc);
            properties.Lines = lineCounter.GetLineCount();
            Assert.AreEqual(14, properties.Lines);

            // Paragraph count: Assign the size of the count of child Paragraph-nodes to the Paragraphs built in property
            properties.Paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Count;
            Assert.AreEqual(2, properties.Paragraphs);

            // Bytes: Use a stream to find out the real file size of our document and assign it to the Property
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Docx);
                properties.Bytes = (int)stream.Length;
                Assert.AreEqual(10871, properties.Bytes);
            }
            
            // Template: The Template attribute can reflect the filename of the attached template document
            doc.AttachedTemplate = MyDir + "Document.BusinessBrochureTemplate.dot";
            Assert.AreEqual("Normal", properties.Template);          
            properties.Template = doc.AttachedTemplate;

            // Content status: This is a descriptive field
            properties.ContentStatus = "Draft";

            // Content type: Upon saving, any value we assign to this field will be overwritten by the MIME type of the output save format
            Assert.AreEqual("", properties.ContentType);

            // If the document contains links and they are all up to date, we can set this to true
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
        public void Thumbnail()
        {
            //ExStart
            //ExFor:Properties.BuiltInDocumentProperties.Thumbnail
            //ExFor:Properties.DocumentProperty.ToByteArray
            //ExSummary:Shows how to append a thumbnail to an Epub document.
            // Create a blank document and add some text with a DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // The thumbnail property resides in a document's built in properties, but is used exclusively by Epub e-book documents
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            // Load an image from our file system into a byte array
            byte[] thumbnailBytes = File.ReadAllBytes(ImageDir + "Aspose.Words.gif");

            // Set the value of the Thumbnail property to the array from above
            properties.Thumbnail = thumbnailBytes;

            // Our thumbnail should be visible at the start of the document, before the text we added
            doc.Save(ArtifactsDir + "Properties.Thumbnail.epub");

            // We can also extract a thumbnail property into a byte array and then into the local file system like this
            DocumentProperty thumbnail = doc.BuiltInDocumentProperties["Thumbnail"];
            File.WriteAllBytes(ArtifactsDir + "Properties.Thumbnail.gif", thumbnail.ToByteArray());
            //ExEnd
        }

        [Test]
        public void HeadingPairs()
        {
            //ExStart
            //ExFor:Properties.BuiltInDocumentProperties.HeadingPairs
            //ExFor:Properties.BuiltInDocumentProperties.TitlesOfParts
            //ExSummary:Shows the relationship between HeadingPairs and TitlesOfParts properties.
            // Open a document that contains entries in the HeadingPairs/TitlesOfParts properties
            Document doc = new Document(MyDir + "Properties.HeadingPairs.docx");

            // The HeadingPairs property is a collection of <string, int> pairs that determines how many document parts a heading spans over
            object[] headingPairs = doc.BuiltInDocumentProperties.HeadingPairs;

            // There are 6 array elements designating 3 heading/part count pairs
            Assert.AreEqual(6, headingPairs.Length);
            Assert.AreEqual("Title", headingPairs[0].ToString());
            Assert.AreEqual("1", headingPairs[1].ToString());
            Assert.AreEqual("Heading 1", headingPairs[2].ToString());
            Assert.AreEqual("5", headingPairs[3].ToString());
            Assert.AreEqual("Heading 2", headingPairs[4].ToString());
            Assert.AreEqual("2", headingPairs[5].ToString());

            // The TitlesOfParts property contains the names of parts that belong to the above headings
            // The above headings preside over (1 + 5 + 2) = 8 parts, and this collection contains their names 
            string[] titlesOfParts = doc.BuiltInDocumentProperties.TitlesOfParts;
            Assert.AreEqual(8, titlesOfParts.Length);

            // "Title"
            Assert.AreEqual("", titlesOfParts[0]);

            // "Heading 1"
            Assert.AreEqual("Part1", titlesOfParts[1]);
            Assert.AreEqual("Part2", titlesOfParts[2]);
            Assert.AreEqual("Part3", titlesOfParts[3]);
            Assert.AreEqual("Part4", titlesOfParts[4]);
            Assert.AreEqual("Part5", titlesOfParts[5]);

            // "Heading 2"
            Assert.AreEqual("Part6", titlesOfParts[6]);
            Assert.AreEqual("Part7", titlesOfParts[7]);

            // We can find the combined values of these collections in File > Properties > Advanced Properties > Contents tab
            // In the special case of the Title property, changing its name like this will change the value of the part name belonging to the "Title" heading
            // By default the name is an empty string, as we can see above, but after saving the document it will inherit the value of the Title property
            doc.BuiltInDocumentProperties.Title = "My Title";
            doc.Save(ArtifactsDir + "Properties.HeadingPairs.docx");
            //ExEnd
        }

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
        public void DocumentPropertyCollection()
        {
            //ExStart
            //ExFor:CustomDocumentProperties.Add(String,String)
            //ExFor:CustomDocumentProperties.Add(String,Boolean)
            //ExFor:CustomDocumentProperties.Add(String,int)
            //ExFor:CustomDocumentProperties.Add(String,DateTime)
            //ExFor:CustomDocumentProperties.Add(String,Double)
            //ExFor:Properties.DocumentPropertyCollection
            //ExFor:Properties.DocumentPropertyCollection.Clear
            //ExFor:Properties.DocumentPropertyCollection.Contains(System.String)
            //ExFor:Properties.DocumentPropertyCollection.GetEnumerator
            //ExFor:Properties.DocumentPropertyCollection.IndexOf(System.String)
            //ExFor:Properties.DocumentPropertyCollection.RemoveAt(System.Int32)
            //ExFor:Properties.DocumentPropertyCollection.Remove
            //ExId:AddCustomProperties
            //ExSummary:Shows how to add custom properties to a document.
            // Create a blank document and get its custom property collection
            Document doc = new Document();
            CustomDocumentProperties properties = doc.CustomDocumentProperties;

            // The collection will be empty by default
            Assert.AreEqual(0, properties.Count);

            // We can populate it with key/value pairs with a variety of value types
            properties.Add("Authorized", true);
            properties.Add("Authorized By", "John Doe");
            properties.Add("Authorized Date", DateTime.Today);
            properties.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
            properties.Add("Authorized Amount", 123.45);

            // Custom properties are automatically sorted in alphabetic order
            Assert.AreEqual(1, properties.IndexOf("Authorized Amount"));
            Assert.AreEqual(5, properties.Count);

            // Enumerate and print all custom properties
            using (IEnumerator<DocumentProperty> enumerator = properties.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"Name: \"{enumerator.Current.Name}\", Type: \"{enumerator.Current.Type}\", Value: \"{enumerator.Current.Value}\"");
                }
            }

            // We can view/edit custom properties by opening the document and looking in File > Properties > Advanced Properties > Custom
            doc.Save(ArtifactsDir + "Properties.DocumentPropertyCollection.docx");

            // We can remove elements from the property collection by index or by name
            properties.RemoveAt(1);
            Assert.False(properties.Contains("Authorized Amount"));
            Assert.AreEqual(4, properties.Count);

            properties.Remove("Authorized Revision");
            Assert.False(properties.Contains("Authorized Revision"));
            Assert.AreEqual(3, properties.Count);

            // We can also empty the entire custom property collection at once
            properties.Clear();
            Assert.AreEqual(0, properties.Count);
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