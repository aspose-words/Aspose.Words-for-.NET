using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Tables;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class LINQtoXMLMailMerge
    {
        public static void Run()
        {
#if !NET20
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 

            // Load the XML document.
            XElement orderXml = XElement.Load(dataDir + "PurchaseOrder.xml");

            // Query the purchase order xml file using LINQ to extract the order items 
            // into an object of an anonymous type. 
            //
            // Make sure you give the properties of the anonymous type the same names as 
            // the MERGEFIELD fields in the document.
            //
            // To pass the actual values stored in the XML element or attribute to Aspose.Words, 
            // we need to cast them to string. This is to prevent the XML tags being inserted into the final document when
            // the XElement or XAttribute objects are passed to Aspose.Words.
            //ExStart:LINQtoXMLMailMergeorderItems
            var orderItems =
            from order in orderXml.Descendants("Item")
            select new
            {
                PartNumber = (string)order.Attribute("PartNumber"),
                ProductName = (string)order.Element("ProductName"),
                Quantity = (string)order.Element("Quantity"),
                USPrice = (string)order.Element("USPrice"),
                Comment = (string)order.Element("Comment"),
                ShipDate = (string)order.Element("ShipDate")
            };
            //ExEnd:LINQtoXMLMailMergeorderItems
            //ExStart:LINQToXMLQueryForDeliveryAddress
            var deliveryAddress =
            from delivery in orderXml.Elements("Address")
            where ((string)delivery.Attribute("Type") == "Shipping")
            select new
            {
                Name = (string)delivery.Element("Name"),
                Country = (string)delivery.Element("Country"),
                Zip = (string)delivery.Element("Zip"),
                State = (string)delivery.Element("State"),
                City = (string)delivery.Element("City"),
                Street = (string)delivery.Element("Street")
            };
            //ExEnd:LINQToXMLQueryForDeliveryAddress
            // Create custom Aspose.Words mail merge data sources based on the LINQ queries.
            MyMailMergeDataSource orderItemsDataSource = new MyMailMergeDataSource(orderItems, "Items");
            MyMailMergeDataSource deliveryDataSource = new MyMailMergeDataSource(deliveryAddress);
            //ExStart:LINQToXMLMailMerge
            string fileName = "TestFile.LINQ.doc";
            // Open the template document.
            Document doc = new Document(dataDir + fileName);

            // Fill the document with data from our data sources.
            // Using mail merge regions for populating the order items table is required
            // because it allows the region to be repeated in the document for each order item.
            doc.MailMerge.ExecuteWithRegions(orderItemsDataSource);

            // The standard mail merge without regions is used for the delivery address.
            doc.MailMerge.Execute(deliveryDataSource);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the output document.
            doc.Save(dataDir);
            //ExEnd:LINQToXMLMailMerge
            Console.WriteLine("\nMail merge performed successfully.\nFile saved at " + dataDir);
#else
            throw new InvalidOperationException("This example requires the .NET Framework v3.5 or above to run." +
                                   " Make sure that the target framework of this project is set to 3.5 or above.");
#endif
        }

        /// <summary>
        /// Aspose.Words does not accept LINQ queries as an input for mail merge directly, 
        /// but provides a generic mechanism which allows mail merges from any data source.
        /// 
        /// This class is a simple implementation of the Aspose.Words custom mail merge data source 
        /// interface that accepts a LINQ query (in fact any IEnumerable object).
        /// Aspose.Words calls this class during the mail merge to retrieve the data.
        /// </summary>
        //ExStart:MyMailMergeDataSource 
        public class MyMailMergeDataSource : IMailMergeDataSource
        //ExEnd:MyMailMergeDataSource 
        {
            /// <summary>
            /// Creates a new instance of a custom mail merge data source.
            /// </summary>
            /// <param name="data">Data returned from a LINQ query.</param>
            //ExStart:MyMailMergeDataSourceConstructor 
            public MyMailMergeDataSource(IEnumerable data)
            {
                mEnumerator = data.GetEnumerator();
            }
            //ExEnd:MyMailMergeDataSourceConstructor 
            /// <summary>
            /// Creates a new instance of a custom mail merge data source, for mail merge with regions.
            /// </summary>
            /// <param name="data">Data returned from a LINQ query.</param>
            /// <param name="tableName">Name of the data source is only used when you perform mail merge with regions. 
            /// If you prefer to use the simple mail merge then use constructor with one parameter.</param>          
            //ExStart:MyMailMergeDataSourceConstructorWithDataTable
            public MyMailMergeDataSource(IEnumerable data, string tableName)
            {
                mEnumerator = data.GetEnumerator();
                mTableName = tableName;
            }
            //ExEnd:MyMailMergeDataSourceConstructorWithDataTable
            /// <summary>
            /// Aspose.Words calls this method to get a value for every data field.
            /// 
            /// This is a simple "generic" implementation of a data source that can work over 
            /// any IEnumerable collection. This implementation assumes that the merge field
            /// name in the document matches the name of a public property on the object
            /// in the collection and uses reflection to get the value of the property.
            /// </summary>
            //ExStart:MyMailMergeDataSourceGetValue
            public bool GetValue(string fieldName, out object fieldValue)
            {
                // Use reflection to get the property by name from the current object.
                object obj = mEnumerator.Current;

                Type curentRecordType = obj.GetType();
                PropertyInfo property = curentRecordType.GetProperty(fieldName);
                if (property != null)
                {
                    fieldValue = property.GetValue(obj, null);
                    return true;
                }

                // Return False to the Aspose.Words mail merge engine to indicate the field was not found.
                fieldValue = null;
                return false;
            }
            //ExEnd:MyMailMergeDataSourceGetValue
            /// <summary>
            /// Moves to the next record in the collection.
            /// </summary>            
            //ExStart:MyMailMergeDataSourceMoveNext
            public bool MoveNext()
            {
                return mEnumerator.MoveNext();
            }
            //ExEnd:MyMailMergeDataSourceMoveNext
            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            //ExStart:MyMailMergeDataSourceTableName
            public string TableName
            {
                get { return mTableName; }
            }
            //ExEnd:MyMailMergeDataSourceTableName
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }           
            private readonly IEnumerator mEnumerator;
            private readonly string mTableName;
        }
        
        
    }
}
