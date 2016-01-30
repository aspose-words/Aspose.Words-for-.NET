// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Border
{
    [TestFixture]
    public class ExBorderCollection : QaTestsBase
    {
        [Test]
        public void GetEnumeratorEx()
        {
            //ExStart
            //ExFor:BorderCollection.GetEnumerator
            //ExSummary:Shows how to use GetEnumerator.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            System.Random rnd = new System.Random();

            var enumerator = borders.GetEnumerator();
            while (enumerator.MoveNext())
            {
                // Do something.
                Aspose.Words.Border b = (Aspose.Words.Border)enumerator.Current;
                b.Color = System.Drawing.Color.FromArgb(rnd.Next(0, 255), rnd.Next(0, 255), rnd.Next(0, 255));
            }

            doc.Save(ExDir + "Document.RandomColourBorder.doc");
            //ExEnd
        }

        [Test]
        public void ClearFormattingEx()
        {
            //ExStart
            //ExFor:BorderCollection.ClearFormatting
            //ExSummary:Shows how to use ClearFormatting.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            borders.ClearFormatting();
            //ExEnd
        }
    }
}