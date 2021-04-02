using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace DocsExamples.Mail_Merge_and_Reporting.Custom_examples
{
    class NestedMailMergeCustom : DocsExamplesBase
    {
        [Test]
        public void CustomMailMerge()
        {
            //ExStart:NestedMailMergeCustom
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(" MERGEFIELD TableStart:Customer");

            builder.Write("Full name:\t");
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write("\nAddress:\t");
            builder.InsertField(" MERGEFIELD Address ");
            builder.Write("\nOrders:\n");

            builder.InsertField(" MERGEFIELD TableStart:Order");

            builder.Write("\tItem name:\t");
            builder.InsertField(" MERGEFIELD Name ");
            builder.Write("\n\tQuantity:\t");
            builder.InsertField(" MERGEFIELD Quantity ");
            builder.InsertParagraph();

            builder.InsertField(" MERGEFIELD TableEnd:Order");

            builder.InsertField(" MERGEFIELD TableEnd:Customer");

            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
            };

            customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
            customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
            customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

            // To be able to mail merge from your data source,
            // it must be wrapped into an object that implements the IMailMergeDataSource interface.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            doc.MailMerge.ExecuteWithRegions(customersDataSource);

            doc.Save(ArtifactsDir + "NestedMailMergeCustom.CustomMailMerge.docx");
            //ExEnd:NestedMailMergeCustom
        }

        /// <summary>
        /// An example of a "data entity" class in your application.
        /// </summary>
        public class Customer
        {
            public Customer(string aFullName, string anAddress)
            {
                FullName = aFullName;
                Address = anAddress;
                Orders = new List<Order>();
            }

            public string FullName { get; set; }
            public string Address { get; set; }
            public List<Order> Orders { get; set; }
        }

        /// <summary>
        /// An example of a child "data entity" class in your application.
        /// </summary>
        public class Order
        {
            public Order(string oName, int oQuantity)
            {
                Name = oName;
                Quantity = oQuantity;
            }

            public string Name { get; set; }
            public int Quantity { get; set; }
        }

        /// <summary>
        /// A custom mail merge data source that you implement to allow Aspose.Words
        /// to mail merge data from your Customer objects into Microsoft Word documents.
        /// </summary>
        public class CustomerMailMergeDataSource : IMailMergeDataSource
        {
            public CustomerMailMergeDataSource(List<Customer> customers)
            {
                mCustomers = customers;

                // When the data source is initialized, it must be positioned before the first record.
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName => "Customer";

            /// <summary>
            /// Aspose.Words calls this method to get a value for every data field.
            /// </summary>
            public bool GetValue(string fieldName, out object fieldValue)
            {
                switch (fieldName)
                {
                    case "FullName":
                        fieldValue = mCustomers[mRecordIndex].FullName;
                        return true;
                    case "Address":
                        fieldValue = mCustomers[mRecordIndex].Address;
                        return true;
                    case "Order":
                        fieldValue = mCustomers[mRecordIndex].Orders;
                        return true;
                    default:
                        fieldValue = null;
                        return false;
                }
            }

            /// <summary>
            /// A standard implementation for moving to a next record in a collection.
            /// </summary>
            public bool MoveNext()
            {
                if (!IsEof)
                    mRecordIndex++;

                return !IsEof;
            }

            //ExStart:GetChildDataSourceExample           
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                switch (tableName)
                {
                    // Get the child collection to merge it with the region provided with tableName variable.
                    case "Order":
                        return new OrderMailMergeDataSource(mCustomers[mRecordIndex].Orders);
                    default:
                        return null;
                }
            }
            //ExEnd:GetChildDataSourceExample

            private bool IsEof => (mRecordIndex >= mCustomers.Count);

            private readonly List<Customer> mCustomers;
            private int mRecordIndex;
        }

        public class OrderMailMergeDataSource : IMailMergeDataSource
        {
            public OrderMailMergeDataSource(List<Order> orders)
            {
                mOrders = orders;

                // When the data source is initialized, it must be positioned before the first record.
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName => "Order";

            /// <summary>
            /// Aspose.Words calls this method to get a value for every data field.
            /// </summary>
            public bool GetValue(string fieldName, out object fieldValue)
            {
                switch (fieldName)
                {
                    case "Name":
                        fieldValue = mOrders[mRecordIndex].Name;
                        return true;
                    case "Quantity":
                        fieldValue = mOrders[mRecordIndex].Quantity;
                        return true;
                    default:
                        fieldValue = null;
                        return false;
                }
            }

            /// <summary>
            /// A standard implementation for moving to a next record in a collection.
            /// </summary>
            public bool MoveNext()
            {
                if (!IsEof)
                    mRecordIndex++;

                return !IsEof;
            }

            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                // Return null because we haven't any child elements for this sort of object.
                return null;
            }

            private bool IsEof => mRecordIndex >= mOrders.Count;

            private readonly List<Order> mOrders;
            private int mRecordIndex;
        }
    }
}