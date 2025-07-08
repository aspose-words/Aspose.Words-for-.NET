// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExWordML2003SaveOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void PrettyFormat(bool prettyFormat)
        {
            //ExStart
            //ExFor:WordML2003SaveOptions
            //ExFor:WordML2003SaveOptions.SaveFormat
            //ExSummary:Shows how to manage output document's raw content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
            // to modify how we save the document to the WordML save format.
            WordML2003SaveOptions options = new WordML2003SaveOptions();

            Assert.AreEqual(SaveFormat.WordML, options.SaveFormat);

            // Set the "PrettyFormat" property to "true" to apply tab character indentation and
            // newlines to make the output document's raw content easier to read.
            // Set the "PrettyFormat" property to "false" to save the document's raw content in one continuous body of the text.
            options.PrettyFormat = prettyFormat;

            doc.Save(ArtifactsDir + "WordML2003SaveOptions.PrettyFormat.xml", options);

            string fileContents = File.ReadAllText(ArtifactsDir + "WordML2003SaveOptions.PrettyFormat.xml");
            string newLine = Environment.NewLine;
            if (prettyFormat)
                Assert.IsTrue(fileContents.Contains(
                    string.Format("<o:DocumentProperties>{0}\t\t", newLine) +
                        string.Format("<o:Revision>1</o:Revision>{0}\t\t", newLine) +
                        string.Format("<o:TotalTime>0</o:TotalTime>{0}\t\t", newLine) +
                        string.Format("<o:Pages>1</o:Pages>{0}\t\t", newLine) +
                        string.Format("<o:Words>0</o:Words>{0}\t\t", newLine) +
                        string.Format("<o:Characters>0</o:Characters>{0}\t\t", newLine) +
                        string.Format("<o:Lines>1</o:Lines>{0}\t\t", newLine) +
                        string.Format("<o:Paragraphs>1</o:Paragraphs>{0}\t\t", newLine) +
                        string.Format("<o:CharactersWithSpaces>0</o:CharactersWithSpaces>{0}\t\t", newLine) +
                        string.Format("<o:Version>11.5606</o:Version>{0}\t", newLine) +
                    "</o:DocumentProperties>"));
            else
                Assert.IsTrue(fileContents.Contains(
                    "<o:DocumentProperties><o:Revision>1</o:Revision><o:TotalTime>0</o:TotalTime><o:Pages>1</o:Pages>" +
                    "<o:Words>0</o:Words><o:Characters>0</o:Characters><o:Lines>1</o:Lines><o:Paragraphs>1</o:Paragraphs>" +
                    "<o:CharactersWithSpaces>0</o:CharactersWithSpaces><o:Version>11.5606</o:Version></o:DocumentProperties>"));
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void MemoryOptimization(bool memoryOptimization)
        {
            //ExStart
            //ExFor:WordML2003SaveOptions
            //ExSummary:Shows how to manage memory optimization.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
            // to modify how we save the document to the WordML save format.
            WordML2003SaveOptions options = new WordML2003SaveOptions();

            // Set the "MemoryOptimization" flag to "true" to decrease memory consumption
            // during the document's saving operation at the cost of a longer saving time.
            // Set the "MemoryOptimization" flag to "false" to save the document normally.
            options.MemoryOptimization = memoryOptimization;

            doc.Save(ArtifactsDir + "WordML2003SaveOptions.MemoryOptimization.xml", options);
            //ExEnd
        }
    }
}
