//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExId:LoadTxt
//ExSummary:Loads a plain text file into an Aspose.Words.Document object.
using System;
using System.IO;
using System.Reflection;
using System.Text;

using Aspose.Words;

namespace LoadTxt
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // This object will help us generate the document.
            DocumentBuilder builder = new DocumentBuilder();

            // You might need to specify a different encoding depending on your plain text files.
            using (StreamReader reader = new StreamReader(dataDir + "LoadTxt.txt", Encoding.UTF8))
            {
                // Read plain text "lines" and convert them into paragraphs in the document.
                string line = null;          
                while((line = reader.ReadLine()) != null)
                {
                    builder.Writeln(line);
                }
            }

            // Save in any Aspose.Words supported format.
            builder.Document.Save(dataDir + "LoadTxt Out.docx");
            builder.Document.Save(dataDir + "LoadTxt Out.html");
        }
    }
}
//ExEnd
