// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
