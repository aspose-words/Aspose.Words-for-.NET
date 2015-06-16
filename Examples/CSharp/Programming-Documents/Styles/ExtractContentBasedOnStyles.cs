﻿//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Drawing;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Diagnostics;

namespace CSharp.Programming_Documents.Working_with_Styles
{
    class ExtractContentBasedOnStyles
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithStyles();

            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Define style names as they are specified in the Word document.
            const string paraStyle = "Heading 1";
            const string runStyle = "Intense Emphasis";

            // Collect paragraphs with defined styles. 
            // Show the number of collected paragraphs and display the text of this paragraphs.
            ArrayList paragraphs = ParagraphsByStyleName(doc, paraStyle);
            Console.WriteLine(string.Format("Paragraphs with \"{0}\" styles ({1}):", paraStyle, paragraphs.Count));
            foreach (Paragraph paragraph in paragraphs)
                Console.Write(paragraph.ToString(SaveFormat.Text));

            // Collect runs with defined styles. 
            // Show the number of collected runs and display the text of this runs.
            ArrayList runs = RunsByStyleName(doc, runStyle);
            Console.WriteLine(string.Format("\nRuns with \"{0}\" styles ({1}):", runStyle, runs.Count));
            foreach (Run run in runs)
                Console.WriteLine(run.Range.Text);

            Console.WriteLine("\nExtracted contents based on styles successfully.");
        }

        public static ArrayList ParagraphsByStyleName(Document doc, string styleName)
        {
            // Create an array to collect paragraphs of the specified style.
            ArrayList paragraphsWithStyle = new ArrayList();
            // Get all paragraphs from the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            // Look through all paragraphs to find those with the specified style.
            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.ParagraphFormat.Style.Name == styleName)
                    paragraphsWithStyle.Add(paragraph);
            }
            return paragraphsWithStyle;
        }
        
        public static ArrayList RunsByStyleName(Document doc, string styleName)
        {
            // Create an array to collect runs of the specified style.
            ArrayList runsWithStyle = new ArrayList();
            // Get all runs from the document.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            // Look through all runs to find those with the specified style.
            foreach (Run run in runs)
            {
                if (run.Font.Style.Name == styleName)
                    runsWithStyle.Add(run);
            }
            return runsWithStyle;
        }
    }
}
