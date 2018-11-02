// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTxtLoadOptions : ApiExampleBase
    {
        [Test]
        public void DetectNumberingWithWhitespaces()
        {
            //ExStart
            //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
            //ExFor:TxtLoadOptions.TrailingSpacesOptions
            //ExFor:TxtLoadOptions.LeadingSpacesOptions
            //ExSummary:Shows how to load plain text as is.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                // If it sets to true Aspose.Words insert additional periods after numbers in the content.
                DetectNumberingWithWhitespaces = false, 
                TrailingSpacesOptions = TxtTrailingSpacesOptions.Preserve,
                LeadingSpacesOptions = TxtLeadingSpacesOptions.Preserve
            };

            Document doc = new Document(MyDir + "TxtLoadOptions.DetectNumberingWithWhitespaces.txt", loadOptions);
            doc.Save(MyDir + @"\Artifacts\TxtLoadOptions.DetectNumberingWithWhitespaces.txt");
            //ExEnd
        }
    }
}