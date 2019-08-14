// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDocSaveOptions : ApiExampleBase
    {
        [Test]
        public void SaveAsDoc()
        {
            //ExStart
            //ExFor:DocSaveOptions
            //ExFor:DocSaveOptions.#ctor
            //ExFor:DocSaveOptions.Password
            //ExFor:DocSaveOptions.SaveFormat
            //ExFor:DocSaveOptions.SaveRoutingSlip
            //ExSummary:Shows how to set save options for classic Microsoft Word document versions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Hello world!");

            // DocSaveOptions only applies to Doc and Dot save formats
            DocSaveOptions options = new DocSaveOptions(SaveFormat.Doc);

            // Set a password with which the document will be encrypted, and which will be required to open it
            options.Password = "MyPassword";

            // If the document contains a routing slip, we can preserve it while saving by setting this flag to true
            options.SaveRoutingSlip = true;

            doc.Save(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc", options);          
            //ExEnd
        }
    }
}