//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Drawing;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;

namespace EnumerateLayoutElements
{
    class Program
    {
        static void Main(string[] args)
        {      
            string dataDir = Path.GetFullPath("../../Data/");
 
            Document doc = new Document(dataDir + "TestFile.docx");

            // This creates an enumerator which is used to "walk" the elements of a rendered document.
            LayoutEnumerator it = new LayoutEnumerator(doc);

            // This sample uses the enumerator to write information about each layout element to the console.
            LayoutInfoWriter.Run(it);

            // This sample adds a border around each layout element and saves each page as a JPEG image to the data directory.
            OutlineLayoutEntitiesRenderer.Run(doc, it, dataDir);      
        }
    }
}
