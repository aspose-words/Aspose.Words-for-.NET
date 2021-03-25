// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Data;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class MailMergeFromXmlUsingDataSet : TestUtil
    {
        [Test]
        public static void MailMergeFromXmlUsingDataSetFeature()
        {
            DataSet customersDs = new DataSet();
            customersDs.ReadXml(MyDir + "Mail merge data - Customers.xml");

            Document doc = new Document(MyDir + "Mail merge destinations - Registration complete.docx");
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            doc.Save(ArtifactsDir + "Mail merge from xml using DataSet - Aspose.Words.docx");
        }
    }
}
