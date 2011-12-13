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

using Aspose.Words;

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
            //ExFor:MailMerge.RemoveEmptyRegions
            //ExId:RemoveEmptyRegions
            //ExSummary:Shows how to remove unmerged mail merge regions from the document.
            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Create a dummy data source containing two empty DataTables which corresponds to the regions in the document.
            DataSet data = new DataSet();
            DataTable suppliers = new DataTable();
            DataTable storeDetails = new DataTable();
            suppliers.TableName = "Suppliers";
            storeDetails.TableName = "StoreDetails";
            data.Tables.Add(suppliers);
            data.Tables.Add(storeDetails);

            // Set the RemoveEmptyRegions to true in order to remove unmerged mail merge regions from the document.
            doc.MailMerge.RemoveEmptyRegions = true;

            // Execute mail merge. It will have no effect as there is no data.
            doc.MailMerge.ExecuteWithRegions(data);

            // Save the output document to disk.
            doc.Save(dataDir + "TestFile.RemoveEmptyRegions Out.doc");
            //ExEnd
        }
    }
}