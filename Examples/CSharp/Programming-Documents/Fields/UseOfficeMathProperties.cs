using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Math;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class UseOfficeMathProperties
    {
        public static void Run()
        {
            // ExStart:SpecifylocaleAtFieldlevel
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            Document doc = new Document(dataDir+ "MathEquations.docx");
            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

            // Gets/sets Office Math display format type which represents whether an equation is displayed inline with the text or displayed on its own line.
            officeMath.DisplayType = OfficeMathDisplayType.Display; // or OfficeMathDisplayType.Inline

            // Gets/sets Office Math justification.
            officeMath.Justification = OfficeMathJustification.Left; // Left justification of Math Paragraph.

            doc.Save(dataDir + "MathEquations_out.docx");
            // ExEnd:SpecifylocaleAtFieldlevel
        }
    }
}
