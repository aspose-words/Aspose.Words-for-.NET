// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Data;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class NestedMailMergeRegions : TestUtil
    {
        [Test]
        public static void NestedMailMergeRegionsFeature()
        {
            DataSet pizzaDs = new DataSet();

            // Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
            // To see examples of how to set up relations manually check the corresponding documentation of this sample
            pizzaDs.ReadXml(MyDir + "Mail merge data - Orders.xml");

            Document doc = new Document(MyDir + "Mail merge destinations - Invoice.docx");

            // Execute the nested mail merge with regions.
            doc.MailMerge.ExecuteWithRegions(pizzaDs);

            doc.Save(ArtifactsDir + "Nested mail merge regions - Aspose.Words.docx");
        }
    }
}
