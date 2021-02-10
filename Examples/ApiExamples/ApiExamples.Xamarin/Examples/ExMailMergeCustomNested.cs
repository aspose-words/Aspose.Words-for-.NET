// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExMailMergeCustomNested : ApiExampleBase
    {
        //ExStart
        //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSource)
        //ExSummary:Shows how to use mail merge regions to execute a nested mail merge.
        [Test] //ExSkip
        public void CustomDataSource()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
            // Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
            // Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
            builder.InsertField(" MERGEFIELD TableStart:Customers");

            // These MERGEFIELDs are inside the mail merge region of the "Customers" table.
            // When we execute the mail merge, this field will receive data from rows in a data source named "Customers".
            builder.Write("Full name:\t");
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write("\nAddress:\t");
            builder.InsertField(" MERGEFIELD Address ");
            builder.Write("\nOrders:\n");

            // Create a second mail merge region inside the outer region for a data source named "Orders".
            // The "Orders" data entries have a many-to-one relationship with the "Customers" data source.
            builder.InsertField(" MERGEFIELD TableStart:Orders");

            builder.Write("\tItem name:\t");
            builder.InsertField(" MERGEFIELD Name ");
            builder.Write("\n\tQuantity:\t");
            builder.InsertField(" MERGEFIELD Quantity ");
            builder.InsertParagraph();

            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // Create related data with names that match those of our mail merge regions.
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
            customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
            customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

            // To mail merge from your data source, we must wrap it into an object that implements the IMailMergeDataSource interface.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);
            
            doc.MailMerge.ExecuteWithRegions(customersDataSource);

            doc.Save(ArtifactsDir + "NestedMailMergeCustom.CustomDataSource.docx");
            TestCustomDataSource(customers, new Document(ArtifactsDir + "NestedMailMergeCustom.CustomDataSource.docx")); //ExSkip
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
                Orders = new OrderList();
            }

            public string FullName { get; set; }
            public string Address { get; set; }
            public OrderList Orders { get; set; }
        }

        /// <summary>
        /// An example of a typed collection that contains your "data" objects.
        /// </summary>
        public class CustomerList : ArrayList
        {
            public new Customer this[int index]
            {
                get { return (Customer) base[index]; }
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
                Name = oName;
                Quantity = oQuantity;
            }

            public string Name { get; set; }
            public int Quantity { get; set; }
        }

        /// <summary>
        /// An example of a typed collection that contains your "data" objects.
        /// </summary>
        public class OrderList : ArrayList
        {
            public new Order this[int index]
            {
                get { return (Order) base[index]; }
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

                // When we initialize the data source, its position must be before the first record.
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName
            {
                get { return "Customers"; }
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
                        // Return "false" to the Aspose.Words mail merge engine to signify
                        // that we could not find a field with this name.
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
                switch (tableName)
                {
                    // Get the child data source, whose name matches the mail merge region that uses its columns.
                    case "Orders":
                        return new OrderMailMergeDataSource(mCustomers[mRecordIndex].Orders);
                    default:
                        return null;
                }
            }

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

                // When we initialize the data source, its position must be before the first record.
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName
            {
                get { return "Orders"; }
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
                        // Return "false" to the Aspose.Words mail merge engine to signify
                        // that we could not find a field with this name.
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

            /// <summary>
            /// Return null because we do not have any child elements for this sort of object.
            /// </summary>
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
        //ExEnd

        private void TestCustomDataSource(CustomerList customers, Document doc)
        {
            List<string[]> mailMergeData = new List<string[]>();

            foreach (Customer customer in customers)
            {
                foreach (Order order in customer.Orders)
                    mailMergeData.Add(new []{ order.Name, order.Quantity.ToString() });
                mailMergeData.Add(new [] {customer.FullName, customer.Address});
            }

            TestUtil.MailMergeMatchesArray(mailMergeData.ToArray(), doc, false);
        }
    }
}