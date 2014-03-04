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

namespace SaveAsMultipageTiffExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            //ExStart
            //ExId:SaveAsMultipageTiff_save
            //ExSummary:Convert document to TIFF.
            // Save the document as multipage TIFF.
            doc.Save(dataDir + "TestFile Out.tiff");
            //ExEnd
            
            //ExStart
            //ExId:SaveAsMultipageTiff_SaveWithOptions
            //ExSummary:Convert to TIFF using customized options        
            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageIndex = 0;
            options.PageCount = 2;
            options.TiffCompression = TiffCompression.Ccitt4;
            options.Resolution = 160;
            
            doc.Save(dataDir + "TestFileWithOptions Out.tiff", options);
            //ExEnd
        }
    }
}