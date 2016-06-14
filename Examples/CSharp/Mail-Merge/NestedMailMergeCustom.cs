using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Data;
using System.Collections;
using Aspose.Words.MailMerging;
namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class NestedMailMergeCustom
    {
        public static void Run()
        {
            //ExStart:NestedMailMergeCustom
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            string fileName = "NestedMailMerge.CustomDataSource.doc";
            // Create some data that we will use in the mail merge.
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            // Create some data for nesting in the mail merge.
            customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
            customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
            customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

            // Open the template document.
            Document doc = new Document(dataDir + fileName);

            // To be able to mail merge from your own data source, it must be wrapped
            // into an object that implements the IMailMergeDataSource interface.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // Now you can pass your data source into Aspose.Words.
            doc.MailMerge.ExecuteWithRegions(customersDataSource);          

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            doc.Save(dataDir);
            //ExEnd:NestedMailMergeCustom

            Console.WriteLine("\nMail merge performed with nested custom data successfully.\nFile saved at " + dataDir);
        }

        /// <summary>
        /// An example of a "data entity" class in your application.
        /// </summary>
        public class Customer
        {
            public Customer(string aFullName, string anAddress)
            {
                mFullName = aFullName;
                mAddress = anAddress;
                mOrders = new OrderList();
            }

            public string FullName
            {
                get { return mFullName; }
                set { mFullName = value; }
            }

            public string Address
            {
                get { return mAddress; }
                set { mAddress = value; }
            }

            public OrderList Orders
            {
                get { return mOrders; }
                set { mOrders = value; }
            }

            private string mFullName;
            private string mAddress;
            private OrderList mOrders;
        }

        /// <summary>
        /// An example of a typed collection that contains your "data" objects.
        /// </summary>
        public class CustomerList : ArrayList
        {
            public new Customer this[int index]
            {
                get { return (Customer)base[index]; }
                set { base[index] = value; }
            }
        }

        /// <summary>
        /// An example of a child "data entity" class in your application.
        /// </summary>
        public class Order
        {
            public Order(string oName, int oQuantity)
            {
                mName = oName;
                mQuantity = oQuantity;
            }

            public string Name
            {
                get { return mName; }
                set { mName = value; }
            }

            public int Quantity
            {
                get { return mQuantity; }
                set { mQuantity = value; }
            }

            private string mName;
            private int mQuantity;
        }

        /// <summary>
        /// An example of a typed collection that contains your "data" objects.
        /// </summary>
        public class OrderList : ArrayList
        {
            public new Order this[int index]
            {
                get { return (Order)base[index]; }
                set { base[index] = value; }
            }
        }

        /// <summary>
        /// A custom mail merge data source that you implement to allow Aspose.Words 
        /// to mail merge data from your Customer objects into Microsoft Word documents.
        /// </summary>
        public class CustomerMailMergeDataSource : IMailMergeDataSource
        {
            public CustomerMailMergeDataSource(CustomerList customers)
            {
                mCustomers = customers;

                // When the data source is initialized, it must be positioned before the first record.
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName
            {
                get { return "Customer"; }
            }

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
                        // A field with this name was not found, 
                        // return false to the Aspose.Words mail merge engine.
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

                return (!IsEof);
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

            private bool IsEof
            {
                get { return (mRecordIndex >= mCustomers.Count); }
            }

            private readonly CustomerList mCustomers;
            private int mRecordIndex;
        }

        public class OrderMailMergeDataSource : IMailMergeDataSource
        {
            public OrderMailMergeDataSource(OrderList orders)
            {
                mOrders = orders;

                // When the data source is initialized, it must be positioned before the first record.
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName
            {
                get { return "Order"; }
            }

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
                        // A field with this name was not found, 
                        // return false to the Aspose.Words mail merge engine.
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

                return (!IsEof);
            }

            // Return null because we haven't any child elements for this sort of object.
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }

            private bool IsEof
            {
                get { return (mRecordIndex >= mOrders.Count); }
            }

            private readonly OrderList mOrders;
            private int mRecordIndex;
        }
    }
}
