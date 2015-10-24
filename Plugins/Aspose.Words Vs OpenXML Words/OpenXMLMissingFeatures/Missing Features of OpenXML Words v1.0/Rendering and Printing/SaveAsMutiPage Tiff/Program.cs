// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SaveAsMutiPage_Tiff
{
    class Program
    {
        static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Open the document.
            Document doc = new Document(dataDir + "SaveAsMutiPageTiff.doc");

            //ExStart
            //ExId:SaveAsMultipageTiff_save
            //ExSummary:Convert document to TIFF.
            // Save the document as multipage TIFF.
            doc.Save(dataDir + "SaveAsMutiPageTiff Out.tiff");
            //ExEnd

            //ExStart
            //ExId:SaveAsMultipageTiff_SaveWithOptions
            //ExSummary:Convert to TIFF using customized options        
            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageIndex = 0;
            options.PageCount = doc.PageCount;
            options.TiffCompression = TiffCompression.Ccitt4;
            options.Resolution = 160;

            doc.Save(dataDir + "TiffFileWithOptions Out.tiff", options);
            //ExEnd
        }
    }
}
