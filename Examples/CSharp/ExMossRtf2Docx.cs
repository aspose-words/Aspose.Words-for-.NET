//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using Aspose.Words;

namespace Examples
{
    public class ExMossRtf2Docx
    {
        //ExStart
        //ExId:MossRtf2Docx
        //ExSummary:Converts an RTF document to OOXML.
        public static void ConvertRtfToDocx(string inFileName, string outFileName)
        {
            // Load an RTF file into Aspose.Words.
            Aspose.Words.Document doc = new Aspose.Words.Document(inFileName);

            // Save the document in the OOXML format.
            doc.Save(outFileName, SaveFormat.Docx);
        }
        //ExEnd
    }
}
