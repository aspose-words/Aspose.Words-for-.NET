using Aspose.Words.Drawing;
using Aspose.Words.Layout;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithRevisionOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            SetShowInBalloons(dataDir);
            SetMeasurementUnit(dataDir);
            SetRevisionBarsPosition(dataDir);
        }

        private static void SetShowInBalloons(string dataDir)
        {
            // ExStart:SetShowInBalloons
            Document doc = new Document(dataDir + "Revisions.docx");

            // Get the RevisionOptions object that controls the appearance of revisions
            RevisionOptions revisionOptions = doc.LayoutOptions.RevisionOptions;

            // Show deletion revisions in balloon
            revisionOptions.ShowInBalloons = ShowInBalloons.FormatAndDelete;

            doc.Save(dataDir + "Revisions.ShowRevisionsInBalloons_out.pdf");
            // ExEnd:SetShowInBalloons
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        private static void SetMeasurementUnit(string dataDir)
        {
            // ExStart:SetMeasurementUnit
            Document doc = new Document(dataDir + "Input.docx");

            // Set Measurement Units to Inches
            doc.LayoutOptions.RevisionOptions.MeasurementUnit = MeasurementUnits.Inches;
            // Show deletion revisions in balloon
            doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.FormatAndDelete;
            // Show Comments
            doc.LayoutOptions.ShowComments = true;

            doc.Save(dataDir + "Revisions.SetMeasurementUnit_out.pdf");
            // ExEnd:SetMeasurementUnit
        }

        private static void SetRevisionBarsPosition(string dataDir)
        {
            // ExStart:SetRevisionBarsPosition
            Document doc = new Document(dataDir + "Input.docx");

            //Renders revision bars on the right side of a page.
            doc.LayoutOptions.RevisionOptions.RevisionBarsPosition = HorizontalAlignment.Right;

            doc.Save(dataDir + "Revisions.SetRevisionBarsPosition_out.pdf");
            // ExEnd:SetRevisionBarsPosition
        }
    }
}
