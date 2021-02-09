using Aspose.Words;
using Aspose.Words.Math;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements
{
    internal class WorkingWithOfficeMath : DocsExamplesBase
    {
        [Test]
        public void MathEquations()
        {
            //ExStart:MathEquations
            Document doc = new Document(MyDir + "Office math.docx");
            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);

            // OfficeMath display type represents whether an equation is displayed inline with the text or displayed on its line.
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;

            doc.Save(ArtifactsDir + "WorkingWithOfficeMath.MathEquations.docx");
            //ExEnd:MathEquations
        }
    }
}