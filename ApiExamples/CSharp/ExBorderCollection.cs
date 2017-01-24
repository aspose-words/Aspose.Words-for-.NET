// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBorderCollection : ApiExampleBase
    {
        [Test]
        public void GetEnumeratorEx()
        {
            //ExStart
            //ExFor:BorderCollection.GetEnumerator
            //ExSummary:Shows how to enumerate all borders in a collection.
            Document doc = new Document(MyDir + "Document.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            var enumerator = borders.GetEnumerator();
            while (enumerator.MoveNext())
            {
                // Do something useful.
                Border b = (Border)enumerator.Current;
                b.Color = Color.RoyalBlue;
                b.LineStyle = LineStyle.Double;
            }

            doc.Save(MyDir + @"\Artifacts\Document.ChangedColourBorder.doc");
            //ExEnd
        }

        [Test]
        public void ClearFormattingEx()
        {
            //ExStart
            //ExFor:BorderCollection.ClearFormatting
            //ExSummary:Shows how to remove all borders from a paragraph at once.
            Document doc = new Document(MyDir + "Document.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            borders.ClearFormatting();
            //ExEnd
        }
    }
}