// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;
using System;
using System.IO;

namespace ApiExamples.LinqTemp
{
    [TestFixture]
    public class ExReportingEngine : ApiExampleBase
    {
        [Test]
        public void BuildingColumnChart()
        {
            //ExStart:BuildingColumnChart
            //GistId:a9bfce4e06620c7bb2f1f0af6d166f0e
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Column Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Column Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Column Chart Report.docx");
            //ExEnd:BuildingColumnChart

            // Test the report.
            CompareDocs("Column Chart Report.docx", "Column Chart Report Gold.docx");
        }

        /// <summary>
        /// A helper method asserting that two files are equal.
        /// </summary>
        /// <param name="artifactFileName">The name of a file within <see cref="ArtifactsDir"/>.</param>
        /// <param name="linqGoldFileName">The name of a file within <see cref="LinqGoldsDir"/>.</param>
        private static void CompareDocs(string artifactFileName, string linqGoldFileName)
        {
            string artifactPath = ArtifactsDir + artifactFileName;
            string linqGoldPath = LinqGoldsDir + linqGoldFileName;

            if (!File.Exists(linqGoldPath))
                File.Copy(artifactPath, linqGoldPath);
            else
                Assert.IsTrue(DocumentHelper.CompareDocs(artifactPath, linqGoldPath));
        }

        /// <summary>
        /// Gets the path to the directory with template and data files used by this class. Ends with a back slash.
        /// </summary>
        private static string MyLinqDir { get; }

        /// <summary>
        /// Gets the path to the directory with previously generated reports used by this class. Ends with a back slash.
        /// </summary>
        private static string LinqGoldsDir { get; }

        static ExReportingEngine()
        {
            MyLinqDir = new Uri(new Uri(MyDir), @"LINQ/").LocalPath;
            LinqGoldsDir = new Uri(new Uri(GoldsDir), @"LINQ/").LocalPath;
        }
    }
}
