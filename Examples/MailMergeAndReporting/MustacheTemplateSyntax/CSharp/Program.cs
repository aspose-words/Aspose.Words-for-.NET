//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Words;
using System.Data;

namespace MustacheTemplateSyntax
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            DataSet ds = new DataSet();

            ds.ReadXml(dataDir + "Orders.xml");

            // Open a template document.
            Document doc = new Document(dataDir + "ExecuteTemplate.doc");

            doc.MailMerge.UseNonMergeFields = true;

            // Execute mail merge to fill the template with data from XML using DataSet.
            doc.MailMerge.ExecuteWithRegions(ds);

            // Save the output document.
            doc.Save(dataDir + "Output.doc");
            
        }
    }
}