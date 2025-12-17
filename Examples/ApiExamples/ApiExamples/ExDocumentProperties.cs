// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Properties;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDocumentProperties : ApiExampleBase
    {
        [Test]
        public void BuiltIn()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties
            //ExFor:Document.BuiltInDocumentProperties
            //ExFor:Document.CustomDocumentProperties
            //ExFor:DocumentProperty
            //ExFor:DocumentProperty.Name
            //ExFor:DocumentProperty.Value
            //ExFor:DocumentProperty.Type
            //ExSummary:Shows how to work with built-in document properties.
            Document doc = new Document(MyDir + "Properties.docx");

            // The "Document" object contains some of its metadata in its members.
            Console.WriteLine($"Document filename:\n\t \"{doc.OriginalFileName}\"");

            // The document also stores metadata in its built-in properties.
            // Each built-in property is a member of the document's "BuiltInDocumentProperties" object.
            Console.WriteLine("Built-in Properties:");
            foreach (DocumentProperty docProperty in doc.BuiltInDocumentProperties)
            {
                Console.WriteLine(docProperty.Name);
                Console.WriteLine($"\tType:\t{docProperty.Type}");

                // Some properties may store multiple values.
                if (docProperty.Value is ICollection<object>)
                {
                    foreach (object value in docProperty.Value as ICollection<object>)
                        Console.WriteLine($"\tValue:\t\"{value}\"");
                }
                else
                {
                    Console.WriteLine($"\tValue:\t\"{docProperty.Value}\"");
                }
            }
            //ExEnd

            Assert.That(doc.BuiltInDocumentProperties.Count, Is.EqualTo(31));
        }

        [Test]
        public void Custom()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.Item(String)
            //ExFor:CustomDocumentProperties
            //ExFor:DocumentProperty.ToString
            //ExFor:DocumentPropertyCollection.Count
            //ExFor:DocumentPropertyCollection.Item(int)
            //ExSummary:Shows how to work with custom document properties.
            Document doc = new Document(MyDir + "Properties.docx");

            // Every document contains a collection of custom properties, which, like the built-in properties, are key-value pairs.
            // The document has a fixed list of built-in properties. The user creates all of the custom properties. 
            Assert.That(doc.CustomDocumentProperties["CustomProperty"].ToString(), Is.EqualTo("Value of custom document property"));

            doc.CustomDocumentProperties.Add("CustomProperty2", "Value of custom document property #2");

            Console.WriteLine("Custom Properties:");
            foreach (var customDocumentProperty in doc.CustomDocumentProperties)
            {
                Console.WriteLine(customDocumentProperty.Name);
                Console.WriteLine($"\tType:\t{customDocumentProperty.Type}");
                Console.WriteLine($"\tValue:\t\"{customDocumentProperty.Value}\"");
            }
            //ExEnd

            Assert.That(doc.CustomDocumentProperties.Count, Is.EqualTo(2));
        }

        [Test]
        public void Description()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.Author
            //ExFor:BuiltInDocumentProperties.Category
            //ExFor:BuiltInDocumentProperties.Comments
            //ExFor:BuiltInDocumentProperties.Keywords
            //ExFor:BuiltInDocumentProperties.Subject
            //ExFor:BuiltInDocumentProperties.Title
            //ExSummary:Shows how to work with built-in document properties in the "Description" category.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            // Below are four built-in document properties that have fields that can display their values in the document body.
            // 1 -  "Author" property, which we can display using an AUTHOR field:
            properties.Author = "John Doe";
            builder.Write("Author:\t");
            builder.InsertField(FieldType.FieldAuthor, true);

            // 2 -  "Title" property, which we can display using a TITLE field:
            properties.Title = "John's Document";
            builder.Write("\nDoc title:\t");
            builder.InsertField(FieldType.FieldTitle, true);

            // 3 -  "Subject" property, which we can display using a SUBJECT field:
            properties.Subject = "My subject";
            builder.Write("\nSubject:\t");
            builder.InsertField(FieldType.FieldSubject, true);

            // 4 -  "Comments" property, which we can display using a COMMENTS field:
            properties.Comments = $"This is {properties.Author}'s document about {properties.Subject}";
            builder.Write("\nComments:\t\"");
            builder.InsertField(FieldType.FieldComments, true);
            builder.Write("\"");

            // The "Category" built-in property does not have a field that can display its value.
            properties.Category = "My category";

            // We can set multiple keywords for a document by separating the string value of the "Keywords" property with semicolons.
            properties.Keywords = "Tag 1; Tag 2; Tag 3";

            // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details".
            // The "Author" built-in property is in the "Origin" group, and the others are in the "Description" group.
            doc.Save(ArtifactsDir + "DocumentProperties.Description.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentProperties.Description.docx");

            properties = doc.BuiltInDocumentProperties;

            Assert.That(properties.Author, Is.EqualTo("John Doe"));
            Assert.That(properties.Category, Is.EqualTo("My category"));
            Assert.That(properties.Comments, Is.EqualTo($"This is {properties.Author}'s document about {properties.Subject}"));
            Assert.That(properties.Keywords, Is.EqualTo("Tag 1; Tag 2; Tag 3"));
            Assert.That(properties.Subject, Is.EqualTo("My subject"));
            Assert.That(properties.Title, Is.EqualTo("John's Document"));
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Author:\t\u0013 AUTHOR \u0014John Doe\u0015\r" +
                            "Doc title:\t\u0013 TITLE \u0014John's Document\u0015\r" +
                            "Subject:\t\u0013 SUBJECT \u0014My subject\u0015\r" +
                            "Comments:\t\"\u0013 COMMENTS \u0014This is John Doe's document about My subject\u0015\""));
        }

        [Test]
        public void Origin()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.Company
            //ExFor:BuiltInDocumentProperties.CreatedTime
            //ExFor:BuiltInDocumentProperties.LastPrinted
            //ExFor:BuiltInDocumentProperties.LastSavedBy
            //ExFor:BuiltInDocumentProperties.LastSavedTime
            //ExFor:BuiltInDocumentProperties.Manager
            //ExFor:BuiltInDocumentProperties.NameOfApplication
            //ExFor:BuiltInDocumentProperties.RevisionNumber
            //ExFor:BuiltInDocumentProperties.Template
            //ExFor:BuiltInDocumentProperties.TotalEditingTime
            //ExFor:BuiltInDocumentProperties.Version
            //ExSummary:Shows how to work with document properties in the "Origin" category.
            // Open a document that we have created and edited using Microsoft Word.
            Document doc = new Document(MyDir + "Properties.docx");
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            // The following built-in properties contain information regarding the creation and editing of this document.
            // We can right-click this document in Windows Explorer and find
            // these properties via "Properties" -> "Details" -> "Origin" category.
            // Fields such as PRINTDATE and EDITTIME can display these values in the document body.
            Console.WriteLine($"Created using {properties.NameOfApplication}, on {properties.CreatedTime}");
            Console.WriteLine($"Minutes spent editing: {properties.TotalEditingTime}");
            Console.WriteLine($"Date/time last printed: {properties.LastPrinted}");
            Console.WriteLine($"Template document: {properties.Template}");

            // We can also change the values of built-in properties.
            properties.Company = "Doe Ltd.";
            properties.Manager = "Jane Doe";
            properties.Version = 5;
            properties.RevisionNumber++;

            // Microsoft Word updates the following properties automatically when we save the document.
            // To use these properties with Aspose.Words, we will need to set values for them manually.
            properties.LastSavedBy = "John Doe";
            properties.LastSavedTime = DateTime.Now;

            // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details" -> "Origin".
            doc.Save(ArtifactsDir + "DocumentProperties.Origin.docx");
            //ExEnd

            properties = new Document(ArtifactsDir + "DocumentProperties.Origin.docx").BuiltInDocumentProperties;

            Assert.That(properties.Company, Is.EqualTo("Doe Ltd."));
            Assert.That(properties.CreatedTime, Is.EqualTo(new DateTime(2006, 4, 25, 10, 10, 0)));
            Assert.That(properties.LastPrinted, Is.EqualTo(new DateTime(2019, 4, 21, 10, 0, 0)));
            Assert.That(properties.LastSavedBy, Is.EqualTo("John Doe"));
            TestUtil.VerifyDate(DateTime.Now, properties.LastSavedTime, TimeSpan.FromSeconds(5));
            Assert.That(properties.Manager, Is.EqualTo("Jane Doe"));
            Assert.That(properties.NameOfApplication, Is.EqualTo("Microsoft Office Word"));
            Assert.That(properties.RevisionNumber, Is.EqualTo(12));
            Assert.That(properties.Template, Is.EqualTo("Normal"));
            Assert.That(properties.TotalEditingTime, Is.EqualTo(8));
            Assert.That(properties.Version, Is.EqualTo(786432));
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
        //ExSummary:Shows how to work with document properties in the "Content" category.
        [Test] //ExSkip
        public void Content()
        {
            Document doc = new Document(MyDir + "Paragraphs.docx");
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            // By using built in properties,
            // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
            // These properties are accessed by right clicking the file in Windows Explorer and navigating to Properties > Details > Content
            // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
            // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
            // Page count: The PageCount property shows the page count in real time and its value can be assigned to the Pages property

            // The "Pages" property stores the page count of the document. 
            Assert.That(properties.Pages, Is.EqualTo(6));

            // The "Words", "Characters", and "CharactersWithSpaces" built-in properties also display various document statistics,
            // but we need to call the "UpdateWordCount" method on the whole document before we can expect them to contain accurate values.
            Assert.That(properties.Words, Is.EqualTo(1054)); //ExSkip
            Assert.That(properties.Characters, Is.EqualTo(6009)); //ExSkip
            Assert.That(properties.CharactersWithSpaces, Is.EqualTo(7049)); //ExSkip
            doc.UpdateWordCount();

            Assert.That(properties.Words, Is.EqualTo(1035));
            Assert.That(properties.Characters, Is.EqualTo(6026));
            Assert.That(properties.CharactersWithSpaces, Is.EqualTo(7041));

            // Count the number of lines in the document, and then assign the result to the "Lines" built-in property.
            LineCounter lineCounter = new LineCounter(doc);
            properties.Lines = lineCounter.GetLineCount();

            Assert.That(properties.Lines, Is.EqualTo(142));

            // Assign the number of Paragraph nodes in the document to the "Paragraphs" built-in property.
            properties.Paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Count;
            Assert.That(properties.Paragraphs, Is.EqualTo(29));

            // Get an estimate of the file size of our document via the "Bytes" built-in property.
            Assert.That(properties.Bytes, Is.EqualTo(20310));

            // Set a different template for our document, and then update the "Template" built-in property manually to reflect this change.
            doc.AttachedTemplate = MyDir + "Business brochure.dotx";

            Assert.That(properties.Template, Is.EqualTo("Normal"));

            properties.Template = doc.AttachedTemplate;

            // "ContentStatus" is a descriptive built-in property.
            properties.ContentStatus = "Draft";

            // Upon saving, the "ContentType" built-in property will contain the MIME type of the output save format.
            Assert.That(properties.ContentType, Is.EqualTo(string.Empty));

            // If the document contains links, and they are all up to date, we can set the "LinksUpToDate" property to "true".
            Assert.That(properties.LinksUpToDate, Is.False);

            doc.Save(ArtifactsDir + "DocumentProperties.Content.docx");
            TestContent(new Document(ArtifactsDir + "DocumentProperties.Content.docx")); //ExSkip
        }

        /// <summary>
        /// Counts the lines in a document.
        /// Traverses the document's layout entities tree upon construction,
        /// counting entities of the "Line" type that also contain real text.
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

        private void TestContent(Document doc)
        {
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            Assert.That(properties.Pages, Is.EqualTo(6));

            Assert.That(properties.Words, Is.EqualTo(1035));
            Assert.That(properties.Characters, Is.EqualTo(6026));
            Assert.That(properties.CharactersWithSpaces, Is.EqualTo(7041));
            Assert.That(properties.Lines, Is.EqualTo(142));
            Assert.That(properties.Paragraphs, Is.EqualTo(29));
            Assert.That(properties.Bytes, Is.EqualTo(15500).Within(200));
            Assert.That(properties.Template, Is.EqualTo(MyDir.Replace("\\\\", "\\") + "Business brochure.dotx"));
            Assert.That(properties.ContentStatus, Is.EqualTo("Draft"));
            Assert.That(properties.ContentType, Is.EqualTo(string.Empty));
            Assert.That(properties.LinksUpToDate, Is.False);
        }

        [Test]
        public void Thumbnail()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.Thumbnail
            //ExFor:DocumentProperty.ToByteArray
            //ExSummary:Shows how to add a thumbnail to a document that we save as an Epub.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // If we save a document, whose "Thumbnail" property contains image data that we added, as an Epub,
            // a reader that opens that document may display the image before the first page.
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;

            byte[] thumbnailBytes = File.ReadAllBytes(ImageDir + "Logo.jpg");
            properties.Thumbnail = thumbnailBytes;

            doc.Save(ArtifactsDir + "DocumentProperties.Thumbnail.epub");

            // We can extract a document's thumbnail image and save it to the local file system.
            DocumentProperty thumbnail = doc.BuiltInDocumentProperties["Thumbnail"];
            File.WriteAllBytes(ArtifactsDir + "DocumentProperties.Thumbnail.gif", thumbnail.ToByteArray());
            //ExEnd

            TestUtil.VerifyImage(400, 400, ArtifactsDir + "DocumentProperties.Thumbnail.gif");
        }

        [Test]
        public void HyperlinkBase()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.HyperlinkBase
            //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a relative hyperlink to a document in the local file system named "Document.docx".
            // Clicking on the link in Microsoft Word will open the designated document, if it is available.
            builder.InsertHyperlink("Relative hyperlink", "Document.docx", false);

            // This link is relative. If there is no "Document.docx" in the same folder
            // as the document that contains this link, the link will be broken.
            Assert.That(File.Exists(ArtifactsDir + "Document.docx"), Is.False);
            doc.Save(ArtifactsDir + "DocumentProperties.HyperlinkBase.BrokenLink.docx");

            // The document we are trying to link to is in a different directory to the one we are planning to save the document in.
            // We could fix links like this by putting an absolute filename in each one. 
            // Alternatively, we could provide a base link that every hyperlink with a relative filename
            // will prepend to its link when we click on it. 
            BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;
            properties.HyperlinkBase = MyDir;

            Assert.That(File.Exists(properties.HyperlinkBase + ((FieldHyperlink)doc.Range.Fields[0]).Address), Is.True);

            doc.Save(ArtifactsDir + "DocumentProperties.HyperlinkBase.WorkingLink.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentProperties.HyperlinkBase.BrokenLink.docx");
            properties = doc.BuiltInDocumentProperties;

            Assert.That(properties.HyperlinkBase, Is.EqualTo(string.Empty));

            doc = new Document(ArtifactsDir + "DocumentProperties.HyperlinkBase.WorkingLink.docx");
            properties = doc.BuiltInDocumentProperties;

            Assert.That(properties.HyperlinkBase, Is.EqualTo(MyDir));
            Assert.That(File.Exists(properties.HyperlinkBase + ((FieldHyperlink)doc.Range.Fields[0]).Address), Is.True);
        }

        [Test]
        public void HeadingPairs()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.HeadingPairs
            //ExFor:BuiltInDocumentProperties.TitlesOfParts
            //ExSummary:Shows the relationship between "HeadingPairs" and "TitlesOfParts" properties.
            Document doc = new Document(MyDir + "Heading pairs and titles of parts.docx");

            // We can find the combined values of these collections via
            // "File" -> "Properties" -> "Advanced Properties" -> "Contents" tab.
            // The HeadingPairs property is a collection of <string, int> pairs that
            // determines how many document parts a heading spans across.
            object[] headingPairs = doc.BuiltInDocumentProperties.HeadingPairs;

            // The TitlesOfParts property contains the names of parts that belong to the above headings.
            string[] titlesOfParts = doc.BuiltInDocumentProperties.TitlesOfParts;

            int headingPairsIndex = 0;
            int titlesOfPartsIndex = 0;
            while (headingPairsIndex < headingPairs.Length)
            {
                Console.WriteLine($"Parts for {headingPairs[headingPairsIndex++]}:");
                int partsCount = Convert.ToInt32(headingPairs[headingPairsIndex++]);

                for (int i = 0; i < partsCount; i++)
                    Console.WriteLine($"\t\"{titlesOfParts[titlesOfPartsIndex++]}\"");
            }
            //ExEnd

            // There are 6 array elements designating 3 heading/part count pairs
            Assert.That(headingPairs.Length, Is.EqualTo(6));
            Assert.That(headingPairs[0].ToString(), Is.EqualTo("Title"));
            Assert.That(headingPairs[1].ToString(), Is.EqualTo("1"));
            Assert.That(headingPairs[2].ToString(), Is.EqualTo("Heading 1"));
            Assert.That(headingPairs[3].ToString(), Is.EqualTo("5"));
            Assert.That(headingPairs[4].ToString(), Is.EqualTo("Heading 2"));
            Assert.That(headingPairs[5].ToString(), Is.EqualTo("2"));

            Assert.That(titlesOfParts.Length, Is.EqualTo(8));
            // "Title"
            Assert.That(titlesOfParts[0], Is.EqualTo(""));
            // "Heading 1"
            Assert.That(titlesOfParts[1], Is.EqualTo("Part1"));
            Assert.That(titlesOfParts[2], Is.EqualTo("Part2"));
            Assert.That(titlesOfParts[3], Is.EqualTo("Part3"));
            Assert.That(titlesOfParts[4], Is.EqualTo("Part4"));
            Assert.That(titlesOfParts[5], Is.EqualTo("Part5"));
            // "Heading 2"
            Assert.That(titlesOfParts[6], Is.EqualTo("Part6"));
            Assert.That(titlesOfParts[7], Is.EqualTo("Part7"));
        }

        [Test]
        public void Security()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.Security
            //ExFor:DocumentSecurity
            //ExSummary:Shows how to use document properties to display the security level of a document.
            Document doc = new Document();

            Assert.That(doc.BuiltInDocumentProperties.Security, Is.EqualTo(DocumentSecurity.None));

            // If we configure a document to be read-only, it will display this status using the "Security" built-in property.
            doc.WriteProtection.ReadOnlyRecommended = true;
            doc.Save(ArtifactsDir + "DocumentProperties.Security.ReadOnlyRecommended.docx");

            Assert.That(new Document(ArtifactsDir + "DocumentProperties.Security.ReadOnlyRecommended.docx").BuiltInDocumentProperties.Security, Is.EqualTo(DocumentSecurity.ReadOnlyRecommended));

            // Write-protect a document, and then verify its security level.
            doc = new Document();

            Assert.That(doc.WriteProtection.IsWriteProtected, Is.False);

            doc.WriteProtection.SetPassword("MyPassword");

            Assert.That(doc.WriteProtection.ValidatePassword("MyPassword"), Is.True);
            Assert.That(doc.WriteProtection.IsWriteProtected, Is.True);

            doc.Save(ArtifactsDir + "DocumentProperties.Security.ReadOnlyEnforced.docx");
            
            Assert.That(new Document(ArtifactsDir + "DocumentProperties.Security.ReadOnlyEnforced.docx").BuiltInDocumentProperties.Security, Is.EqualTo(DocumentSecurity.ReadOnlyEnforced));

            // "Security" is a descriptive property. We can edit its value manually.
            doc = new Document();

            doc.Protect(ProtectionType.AllowOnlyComments, "MyPassword");
            doc.BuiltInDocumentProperties.Security = DocumentSecurity.ReadOnlyExceptAnnotations;
            doc.Save(ArtifactsDir + "DocumentProperties.Security.ReadOnlyExceptAnnotations.docx");

            Assert.That(new Document(ArtifactsDir + "DocumentProperties.Security.ReadOnlyExceptAnnotations.docx").BuiltInDocumentProperties.Security, Is.EqualTo(DocumentSecurity.ReadOnlyExceptAnnotations));
            //ExEnd
        }

        [Test]
        public void CustomNamedAccess()
        {
            //ExStart
            //ExFor:DocumentPropertyCollection.Item(String)
            //ExFor:CustomDocumentProperties.Add(String,DateTime)
            //ExFor:DocumentProperty.ToDateTime
            //ExSummary:Shows how to create a custom document property which contains a date and time.
            Document doc = new Document();

            doc.CustomDocumentProperties.Add("AuthorizationDate", DateTime.Now);
            DateTime authorizationDate = doc.CustomDocumentProperties["AuthorizationDate"].ToDateTime();
            Console.WriteLine($"Document authorized on {authorizationDate}");
            //ExEnd

            TestUtil.VerifyDate(DateTime.Now, 
                DocumentHelper.SaveOpen(doc).CustomDocumentProperties["AuthorizationDate"].ToDateTime(), 
                TimeSpan.FromSeconds(1));
        }

        [Test]
        public void LinkCustomDocumentPropertiesToBookmark()
        {
            //ExStart
            //ExFor:CustomDocumentProperties.AddLinkToContent(String, String)
            //ExFor:DocumentProperty.IsLinkToContent
            //ExFor:DocumentProperty.LinkSource
            //ExSummary:Shows how to link a custom document property to a bookmark.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");
            builder.Write("Hello world!");
            builder.EndBookmark("MyBookmark");

            // Link a new custom property to a bookmark. The value of this property
            // will be the contents of the bookmark that it references in the "LinkSource" member.
            CustomDocumentProperties customProperties = doc.CustomDocumentProperties;
            DocumentProperty customProperty = customProperties.AddLinkToContent("Bookmark", "MyBookmark");

            Assert.That(customProperty.IsLinkToContent, Is.EqualTo(true));
            Assert.That(customProperty.LinkSource, Is.EqualTo("MyBookmark"));
            Assert.That(customProperty.Value, Is.EqualTo("Hello world!"));

            doc.Save(ArtifactsDir + "DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
            customProperty = doc.CustomDocumentProperties["Bookmark"];

            Assert.That(customProperty.IsLinkToContent, Is.EqualTo(true));
            Assert.That(customProperty.LinkSource, Is.EqualTo("MyBookmark"));
            Assert.That(customProperty.Value, Is.EqualTo("Hello world!"));
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
            //ExFor:DocumentProperty.Type
            //ExFor:DocumentPropertyCollection
            //ExFor:DocumentPropertyCollection.Clear
            //ExFor:DocumentPropertyCollection.Contains(String)
            //ExFor:DocumentPropertyCollection.GetEnumerator
            //ExFor:DocumentPropertyCollection.IndexOf(String)
            //ExFor:DocumentPropertyCollection.RemoveAt(Int32)
            //ExFor:DocumentPropertyCollection.Remove
            //ExFor:PropertyType
            //ExSummary:Shows how to work with a document's custom properties.
            Document doc = new Document();
            CustomDocumentProperties properties = doc.CustomDocumentProperties;

            Assert.That(properties.Count, Is.EqualTo(0));

            // Custom document properties are key-value pairs that we can add to the document.
            properties.Add("Authorized", true);
            properties.Add("Authorized By", "John Doe");
            properties.Add("Authorized Date", DateTime.Today);
            properties.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
            properties.Add("Authorized Amount", 123.45);

            // The collection sorts the custom properties in alphabetic order.
            Assert.That(properties.IndexOf("Authorized Amount"), Is.EqualTo(1));
            Assert.That(properties.Count, Is.EqualTo(5));

            // Print every custom property in the document.
            using (IEnumerator<DocumentProperty> enumerator = properties.GetEnumerator())
            {
                while (enumerator.MoveNext())
                    Console.WriteLine($"Name: \"{enumerator.Current.Name}\"\n\tType: \"{enumerator.Current.Type}\"\n\tValue: \"{enumerator.Current.Value}\"");
            }

            // Display the value of a custom property using a DOCPROPERTY field.
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldDocProperty field = (FieldDocProperty)builder.InsertField(" DOCPROPERTY \"Authorized By\"");
            field.Update();

            Assert.That(field.Result, Is.EqualTo("John Doe"));

            // We can find these custom properties in Microsoft Word via "File" -> "Properties" > "Advanced Properties" > "Custom".
            doc.Save(ArtifactsDir + "DocumentProperties.DocumentPropertyCollection.docx");

            // Below are three ways or removing custom properties from a document.
            // 1 -  Remove by index:
            properties.RemoveAt(1);

            Assert.That(properties.Contains("Authorized Amount"), Is.False);
            Assert.That(properties.Count, Is.EqualTo(4));

            // 2 -  Remove by name:
            properties.Remove("Authorized Revision");

            Assert.That(properties.Contains("Authorized Revision"), Is.False);
            Assert.That(properties.Count, Is.EqualTo(3));

            // 3 -  Empty the entire collection at once:
            properties.Clear();

            Assert.That(properties.Count, Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void PropertyTypes()
        {
            //ExStart
            //ExFor:DocumentProperty.ToBool
            //ExFor:DocumentProperty.ToInt
            //ExFor:DocumentProperty.ToDouble
            //ExFor:DocumentProperty.ToString
            //ExFor:DocumentProperty.ToDateTime
            //ExSummary:Shows various type conversion methods of custom document properties.
            Document doc = new Document();
            CustomDocumentProperties properties = doc.CustomDocumentProperties;

            DateTime authDate = DateTime.Today;
            properties.Add("Authorized", true);
            properties.Add("Authorized By", "John Doe");
            properties.Add("Authorized Date", authDate);
            properties.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
            properties.Add("Authorized Amount", 123.45);

            Assert.That(properties["Authorized"].ToBool(), Is.EqualTo(true));
            Assert.That(properties["Authorized By"].ToString(), Is.EqualTo("John Doe"));
            Assert.That(properties["Authorized Date"].ToDateTime(), Is.EqualTo(authDate));
            Assert.That(properties["Authorized Revision"].ToInt(), Is.EqualTo(1));
            Assert.That(properties["Authorized Amount"].ToDouble(), Is.EqualTo(123.45d));
            //ExEnd
        }

        [Test]
        public void ExtendedProperties()
        {
            //ExStart:ExtendedProperties
            //GistId:366eb64fd56dec3c2eaa40410e594182
            //ExFor:BuiltInDocumentProperties.ScaleCrop
            //ExFor:BuiltInDocumentProperties.SharedDocument
            //ExFor:BuiltInDocumentProperties.HyperlinksChanged
            //ExSummary:Shows how to get extended properties.
            Document doc = new Document(MyDir + "Extended properties.docx");
            Assert.That(doc.BuiltInDocumentProperties.ScaleCrop, Is.True);
            Assert.That(doc.BuiltInDocumentProperties.SharedDocument, Is.True);
            Assert.That(doc.BuiltInDocumentProperties.HyperlinksChanged, Is.True);
            //ExEnd:ExtendedProperties
        }
    }
}