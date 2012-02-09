//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;
using System.Data;
using System.Diagnostics;

using Aspose.Words;
using Aspose.Words.Reporting;

namespace RemoveEmptyRegions
{
    class RemoveEmptyRegions
    {
        public static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            //ExStart
            //ExId:RemoveEmptyRegions
            //ExSummary:Shows how to remove unmerged mail merge regions from the document.
            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Create a dummy data source containing no data.
            DataSet data = new DataSet();

            // Set the appropriate mail merge clean up options to remove any unused regions from the document.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;

            // Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
            // automatically as they are unused.
            doc.MailMerge.ExecuteWithRegions(data);

            // Save the output document to disk.
            doc.Save(dataDir + "TestFile.RemoveEmptyRegions Out.doc");
            //ExEnd

            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0, "Error: There are still unused regions remaining in the document");
        }
    }
}