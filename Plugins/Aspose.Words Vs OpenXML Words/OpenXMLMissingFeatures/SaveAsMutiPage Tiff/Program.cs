// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.IO;
using System.Reflection;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
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
