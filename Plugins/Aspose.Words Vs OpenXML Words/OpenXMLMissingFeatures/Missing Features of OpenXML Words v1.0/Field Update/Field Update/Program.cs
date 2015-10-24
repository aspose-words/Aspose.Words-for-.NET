// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Field_Update
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"F:\Dropbox\Personal\Aspose Vs OpenXML\Features Supported by Aspose not Open XML\Aspose.Words Features\Field Update\Data\";
            Document doc = new Document(MyDir + "Rendering.docx");

            // This updates all fields in the document.
            doc.UpdateFields();

            doc.Save(MyDir + "Rendering.UpdateFields Out.pdf");
        }
    }
}
