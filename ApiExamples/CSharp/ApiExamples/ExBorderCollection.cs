// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBorderCollection : ApiExampleBase
    {
        [Test]
        public void GetBordersEnumerator()
        {
            //ExStart
            //ExFor:BorderCollection.GetEnumerator
            //ExSummary:Shows how to enumerate all borders in a collection.
            Document doc = new Document(MyDir + "Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            BorderCollection borders = builder.ParagraphFormat.Borders;

            using (IEnumerator<Border> enumerator = borders.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    // Do something useful.
                    Border b = enumerator.Current;
                    b.Color = Color.RoyalBlue;
                    b.LineStyle = LineStyle.Double;
                }
            }

            doc.Save(ArtifactsDir + "Border.ChangedColorBorder.doc");                                                    
            //ExEnd
        }

        [Test]
        public void RemoveAllBorders()
        {
            //ExStart
            //ExFor:BorderCollection.ClearFormatting
            //ExSummary:Shows how to remove all borders from a paragraph at once.
            Document doc = new Document(MyDir + "Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            borders.ClearFormatting();
            //ExEnd
        }
    }
}