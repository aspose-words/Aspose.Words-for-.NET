// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Create the Dataset and read the XML.
            DataSet pizzaDs = new DataSet();

            // Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
            // To see examples of how to set up relations manually check the corresponding documentation of this sample
            pizzaDs.ReadXml(dataDir + "CustomerData.xml");

            // Open the template document.
            Document doc = new Document(dataDir + "Invoice Template.doc");

            // Execute the nested mail merge with regions
            doc.MailMerge.ExecuteWithRegions(pizzaDs);

            // Save the output to file
            doc.Save(dataDir + "Invoice Out.doc");

            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0, "There was a problem with mail merge");
        }
    }
    public class DataRelationExample
    {
        public static void CreateRelationship()
        {
            DataSet dataSet = new DataSet();
            DataTable orderTable = new DataTable();
            DataTable itemTable = new DataTable();
            //ExStart
            //ExId:NestedMailMergeCreateRelationship
            //ExSummary:Shows how to create a simple DataRelation for use in nested mail merge.
            dataSet.Relations.Add(new DataRelation("OrderToItem", orderTable.Columns["Order_Id"], itemTable.Columns["Order_Id"]));
            //ExEnd
        }

        public static void DisableForeignKeyConstraints()
        {
            DataSet dataSet = new DataSet();
            DataTable orderTable = new DataTable();
            DataTable itemTable = new DataTable();
            //ExStart
            //ExId:NestedMailMergeDisableConstraints
            //ExSummary:Shows how to disable foreign key constraints when creating a DataRelation for use in nested mail merge.
            dataSet.Relations.Add(new DataRelation("OrderToItem", orderTable.Columns["Order_Id"], itemTable.Columns["Order_Id"], false));
            //ExEnd
        }
    }
}
