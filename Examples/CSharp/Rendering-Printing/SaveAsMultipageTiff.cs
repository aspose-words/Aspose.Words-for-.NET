//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;

using Aspose.Words;
using Aspose.Words.Saving;

namespace CSharp.Rendering_and_Printing
{
    class SaveAsMultipageTiff
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_RenderingAndPrinting(); ;

            // Open the document.
            Document doc = new Document(dataDir + "TestFile Multipage TIFF.doc");

            // Save the document as multipage TIFF.
            doc.Save(dataDir + "TestFile Multipage TIFF Out.tiff");
            
            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageIndex = 0;
            options.PageCount = 2;
            options.TiffCompression = TiffCompression.Ccitt4;
            options.Resolution = 160;

            doc.Save(dataDir + "TestFileWithOptions Out.tiff", options);

            Console.WriteLine("\nDocument saved as multi-page TIFF successfully.\nFile saved at " + dataDir + "TestFileWithOptions Out.tiff");
        }
    }
}
