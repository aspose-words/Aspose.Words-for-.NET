//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;

using Aspose.Words;
using Aspose.Words.Drawing;

namespace CompressImages
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;
            string srcFileName = dataDir + "Test.docx";

            Console.WriteLine("Loading {0}. Size {1}.", srcFileName, GetFileSize(srcFileName));
            Document doc = new Document(srcFileName);

            // 220ppi Print - said to be excellent on most printers and screens.
            // 150ppi Screen - said to be good for web pages and projectors.
            // 96ppi Email - said to be good for minimal document size and sharing.
            const int desiredPpi = 150;

            // In .NET this seems to be a good compression / quality setting.
            const int jpegQuality = 90;

            // Resample images to desired ppi and save.
            int count = Resampler.Resample(doc, desiredPpi, jpegQuality);

            Console.WriteLine("Resampled {0} images.", count);

            if (count != 1)
                Console.WriteLine("We expected to have only 1 image resampled in this test document!");

            string dstFileName = srcFileName + ".Resampled Out.docx";
            doc.Save(dstFileName);
            Console.WriteLine("Saving {0}. Size {1}.", dstFileName, GetFileSize(dstFileName));

            // Verify that the first image was compressed by checking the new Ppi.
            doc = new Document(dstFileName);
            DrawingML shape = (DrawingML)doc.GetChild(NodeType.DrawingML, 0, true);
            double imagePpi = shape.ImageData.ImageSize.WidthPixels / ConvertUtil.PointToInch(shape.Size.Width);

            Debug.Assert(imagePpi < 150, "Image was not resampled successfully.");

            Console.WriteLine("Press any key.");
            Console.ReadLine();
        }
        public static int GetFileSize(string fileName)
        {
            using (Stream stream = File.OpenRead(fileName))
                return (int)stream.Length;
        }
    }
}
