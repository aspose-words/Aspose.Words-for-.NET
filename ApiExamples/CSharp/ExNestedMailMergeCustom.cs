// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections;

using Aspose.Words;
using Aspose.Words.MailMerging;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExNestedMailMergeCustom : ApiExampleBase
    {
        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void MailMergeCustomDataSourceCaller()
        {
            this.MailMergeCustomDataSource();
        }

        public void MailMergeCustomDataSource()
        {
            // Create some data that we will use in the mail merge.
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            // Create some data for nesting in the mail merge.
            customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
            customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
            customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

            // Open the template document.
            Document doc = new Document(MyDir + "NestedMailMerge.CustomDataSource.doc");

            // To be able to mail merge from your own data source, it must be wrapped
            // into an object that implements the IMailMergeDataSource interface.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // Now you can pass your data source into Aspose.Words.
            doc.MailMerge.ExecuteWithRegions(customersDataSource);

            doc.Save(MyDir + @"\Artifacts\NestedMailMerge.CustomDataSource.doc");
        }

        /// <summary>
        /// An example of a "data entity" class in your application.
        /// </summary>
        public class Customer
        {
            public Customer(string aFullName, string anAddress)
            {
                this.mFullName = aFullName;
                this.mAddress = anAddress;
                this.mOrders = new OrderList();
            }

            public string FullName
            {
                get { return this.mFullName; }
                set { this.mFullName = value; }
            }

            public string Address
            {
                get { return this.mAddress; }
                set { this.mAddress = value; }
            }

            public OrderList Orders
            {
                get { return this.mOrders; }
                set { this.mOrders = value; }
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
                this.mName = oName;
                this.mQuantity = oQuantity;
            }

            public string Name
            {
                get { return this.mName; }
                set { this.mName = value; }
            }

            public int Quantity
            {
                get { return this.mQuantity; }
                set { this.mQuantity = value; }
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
                this.mCustomers = customers;

                // When the data source is initialized, it must be positioned before the first record.
                this.mRecordIndex= -1;
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
                        fieldValue = this.mCustomers[this.mRecordIndex].FullName;
                        return true;
                    case "Address":
                        fieldValue = this.mCustomers[this.mRecordIndex].Address;
                        return true;
                    case "Order":
                        fieldValue = this.mCustomers[this.mRecordIndex].Orders;
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
                if (!this.IsEof)
                    this.mRecordIndex++;

                return (!this.IsEof);
            }

            //ExStart
            //ExId:GetChildDataSourceExample
            //ExSummary:Shows how to get a child collection of objects by using the GetChildDataSource method in the parent class.
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                switch (tableName)
                {
                    // Get the child collection to merge it with the region provided with tableName variable.
                    case "Order":
                        return new OrderMailMergeDataSource(this.mCustomers[this.mRecordIndex].Orders);
                    default:
                        return null;
                }
            }
            //ExEnd

            private bool IsEof
            {
                get { return (this.mRecordIndex >= this.mCustomers.Count); }
            }

            private readonly CustomerList mCustomers;
            private int mRecordIndex;
        }

        public class OrderMailMergeDataSource : IMailMergeDataSource
        {
            public OrderMailMergeDataSource(OrderList orders)
            {
                this.mOrders = orders;

                // When the data source is initialized, it must be positioned before the first record.
                this.mRecordIndex = -1;
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
                        fieldValue = this.mOrders[this.mRecordIndex].Name;
                        return true;
                    case "Quantity":
                        fieldValue = this.mOrders[this.mRecordIndex].Quantity;
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
                if (!this.IsEof)
                    this.mRecordIndex++;

                return (!this.IsEof);
            }

            // Return null because we haven't any child elements for this sort of object.
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }

            private bool IsEof
            {
                get { return (this.mRecordIndex >= this.mOrders.Count); }
            }

            private readonly OrderList mOrders;
            private int mRecordIndex;
        }
    }
}
