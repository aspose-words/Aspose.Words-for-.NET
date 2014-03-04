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
using System.Text;

using Aspose.Words;

namespace LoadTxtExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // The encoding of the text file is automatically detected.
            Document doc = new Document(dataDir + "LoadTxt.txt");

            // Save as any Aspose.Words supported format, such as DOCX.
            doc.Save(dataDir + "LoadTxt Out.docx");
        }
    }
}
//ExEnd