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
    public class ExMailMergeCustom : ApiExampleBase
    {
        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void MailMergeCustomDataSourceCaller()
        {
            this.MailMergeCustomDataSource();
        }

        //ExStart
        //ExFor:IMailMergeDataSource
        //ExFor:IMailMergeDataSource.TableName
        //ExFor:IMailMergeDataSource.MoveNext
        //ExFor:IMailMergeDataSource.GetValue
        //ExFor:IMailMergeDataSource.GetChildDataSource
        //ExFor:MailMerge.Execute(IMailMergeDataSource)
        //ExSummary:Performs mail merge from a custom data source.
        public void MailMergeCustomDataSource()
        {
            // Create some data that we will use in the mail merge.
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            // Open the template document.
            Document doc = new Document(MyDir + "MailMerge.CustomDataSource.doc");

            // To be able to mail merge from your own data source, it must be wrapped
            // into an object that implements the IMailMergeDataSource interface.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // Now you can pass your data source into Aspose.Words.
            doc.MailMerge.Execute(customersDataSource);

            doc.Save(MyDir + @"\Artifacts\MailMerge.CustomDataSource.doc");
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

            private string mFullName;
            private string mAddress;
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

            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }

            private bool IsEof
            {
                get { return (this.mRecordIndex >= this.mCustomers.Count); }
            }

            private readonly CustomerList mCustomers;
            private int mRecordIndex;
        }
        //ExEnd
    }
}
