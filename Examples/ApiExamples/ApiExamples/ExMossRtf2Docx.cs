// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;

namespace ApiExamples
{
    public static class ExMossRtf2Docx
    {
        public static void ConvertRtfToDocx(string inFileName, string outFileName)
        {
            // Load an RTF file into Aspose.Words.
            Document doc = new Document(inFileName);

            // Save the document in the OOXML format.
            doc.Save(outFileName, SaveFormat.Docx);
        }
    }
}