// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
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
            //ExFor:DocSaveOptions.#ctor(SaveFormat)
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

        [Test]
        public void TempFolder()
        {
            //ExStart
            //ExFor:SaveOptions.TempFolder
            //ExSummary:Shows how to save a document using temporary files.
            Document doc = new Document(MyDir + "Rendering.doc");

            // We can use a SaveOptions object to set the saving method of a document from a MemoryStream to temporary files
            // While saving, the files will briefly pop up in the folder we set as the TempFolder attribute below
            // Doing this will free up space in the memory that the stream would usually occupy
            DocSaveOptions options = new DocSaveOptions();
            options.TempFolder = ArtifactsDir + "TempFiles";

            // Ensure that the directory exists and save
            Directory.CreateDirectory(options.TempFolder);

            doc.Save(ArtifactsDir + "DocSaveOptions.TempFolder.doc", options);
            //ExEnd
        }
    }
}