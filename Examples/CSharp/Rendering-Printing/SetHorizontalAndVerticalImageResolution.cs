using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class SetHorizontalAndVerticalImageResolution
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            // ExStart:SetHorizontalAndVerticalImageResolution
            // Load the documents 
            Document doc = new Document(dataDir + "TestFile.doc");

            //Renders a page of a Word document into a PNG image at a specific horizontal and vertical resolution.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.HorizontalResolution = 300;
            options.VerticalResolution = 300;
            options.PageCount = 1;

            doc.Save(dataDir + "Rendering.SaveToImageResolution Out.png", options);
            // ExEnd:SetHorizontalAndVerticalImageResolution
        }
    }
}
