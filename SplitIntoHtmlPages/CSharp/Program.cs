//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;

namespace SplitIntoHtmlPages
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // You need to have a valid license for Aspose.Words.
            // The best way is to embed the license as a resource into the project
            // and specify only file name without path in the following call.
            // Aspose.Words.License license = new Aspose.Words.License();
            // license.SetLicense(@"Aspose.Words.lic");


            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            string srcFileName = dataDir + "SOI 2007-2012-DeeM with footnote added.doc";
            string tocTemplate = dataDir + "TocTemplate.doc";

            string outDir = Path.Combine(dataDir, "Out");
            Directory.CreateDirectory(outDir);

            // This class does the job.
            Worker w = new Worker();
            w.Execute(srcFileName, tocTemplate, outDir);

            Console.WriteLine("Success.");
        }
    }
}
