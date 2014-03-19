//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            Document doc = new Document(dataDir + "RemoveField.doc");

            //ExStart
            //ExFor:Field.Remove
            //ExId:DocumentBuilder_RemoveField
            //ExSummary:Removes a field from the document.
            Field field = doc.Range.Fields[0];
            // Calling this method completely removes the field from the document.
            field.Remove();
            //ExEnd        
        }
    }
}