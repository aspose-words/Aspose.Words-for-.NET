﻿// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Style
{
    [TestFixture]
    public class ExStyles : QaTestsBase
    {
        [Test]
        public void GetStyles()
        {
            //ExStart
            //ExFor:DocumentBase.Styles
            //ExFor:Style.Name
            //ExId:GetStyles
            //ExSummary:Shows how to get access to the collection of styles defined in the document.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            StyleCollection styles = doc.Styles;

            foreach (Aspose.Words.Style style in styles)
                Console.WriteLine(style.Name);
            //ExEnd
        }

        [Test]
        public void SetAllStyles()
        {
            //ExStart
            //ExFor:Style.Font
            //ExFor:Style
            //ExSummary:Shows how to change the font formatting of all styles in a document.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            foreach (Aspose.Words.Style style in doc.Styles)
            {
                if (style.Font != null)
                {
                    style.Font.ClearFormatting();
                    style.Font.Size = 20;
                    style.Font.Name = "Arial";
                }
            }
            //ExEnd
        }

        [Test]
        public void ChangeStyleOfTOCLevel()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document();
            //ExStart
            //ExId:ChangeTOCStyle
            //ExSummary:Changes a formatting property used in the first level TOC style.
            // Retrieve the style used for the first level of the TOC and change the formatting of the style.
            doc.Styles[StyleIdentifier.Toc1].Font.Bold = true;
            //ExEnd
        }

        [Test]
        public void ChangeTOCTabStops()
        {
            //ExStart
            //ExFor:TabStop
            //ExFor:ParagraphFormat.TabStops
            //ExFor:Style.StyleIdentifier
            //ExFor:TabStopCollection.RemoveByPosition
            //ExFor:TabStop.Alignment
            //ExFor:TabStop.Position
            //ExFor:TabStop.Leader
            //ExId:ChangeTOCTabStops
            //ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.TableOfContents.doc");

            // Iterate through all paragraphs in the document
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9.
                if (para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 && para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9)
                {
                    // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                    TabStop tab = para.ParagraphFormat.TabStops[0];
                    // Remove the old tab from the collection.
                    para.ParagraphFormat.TabStops.RemoveByPosition(tab.Position);
                    // Insert a new tab using the same properties but at a modified position. 
                    // We could also change the separators used (dots) by passing a different Leader type
                    para.ParagraphFormat.TabStops.Add(tab.Position - 50, tab.Alignment, tab.Leader);
                }
            }

            doc.Save(MyDir + "Document.TableOfContentsTabStops Out.doc");
            //ExEnd
        }

        [Test]
        public void CopyStyleSameDocument()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExFor:Style.Name
            //ExSummary:Demonstrates how to copy a style within the same document.
            // The AddCopy method creates a copy of the specified style and automatically generates a new name for the style, such as "Heading 1_0".
            Aspose.Words.Style newStyle = doc.Styles.AddCopy(doc.Styles["Heading 1"]);

            // You can change the new style name if required as the Style.Name property is read-write.
            newStyle.Name = "My Heading 1";
            //ExEnd

            Assert.NotNull(newStyle);
            Assert.AreEqual("My Heading 1", newStyle.Name);
            Assert.AreEqual(doc.Styles["Heading 1"].Type, newStyle.Type);
        }

        [Test]
        public void CopyStyleDifferentDocument()
        {
            Aspose.Words.Document dstDoc = new Aspose.Words.Document();
            Aspose.Words.Document srcDoc = new Aspose.Words.Document();

            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExSummary:Demonstrates how to copy style from one document into a different document.
            // This is the style in the source document to copy to the destination document.
            Aspose.Words.Style srcStyle = srcDoc.Styles[StyleIdentifier.Heading1];

            // Change the font of the heading style to red.
            srcStyle.Font.Color = Color.Red;

            // The AddCopy method can be used to copy a style from a different document.
            Aspose.Words.Style newStyle = dstDoc.Styles.AddCopy(srcStyle);
            //ExEnd

            Assert.NotNull(newStyle);
            Assert.AreEqual("Heading 1", newStyle.Name);
            Assert.AreEqual(Color.Red.ToArgb(), newStyle.Font.Color.ToArgb());
        }

        [Test]
        public void OverwriteStyleDifferentDocument()
        {         
            Aspose.Words.Document dstDoc = new Aspose.Words.Document();
            Aspose.Words.Document srcDoc = new Aspose.Words.Document();

            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExId:OverwriteStyleDifferentDocument   
            //ExSummary:Demonstrates how to copy a style from one document to another and overide an existing style in the destination document.
            // This is the style in the source document to copy to the destination document.
            Aspose.Words.Style srcStyle = srcDoc.Styles[StyleIdentifier.Heading1];

            // Change the font of the heading style to red.
            srcStyle.Font.Color = Color.Red;

            // The AddCopy method can be used to copy a style to a different document.
            Aspose.Words.Style newStyle = dstDoc.Styles.AddCopy(srcStyle);

            // The name of the new style can be changed to the name of any existing style. Doing this will override the existing style.
            newStyle.Name = "Heading 1";
            //ExEnd

            Assert.NotNull(newStyle);
            Assert.AreEqual("Heading 1", newStyle.Name);
            Assert.IsNull(dstDoc.Styles["Heading 1_0"]);
            Assert.AreEqual(Color.Red.ToArgb(), newStyle.Font.Color.ToArgb());
        }
    }
}
