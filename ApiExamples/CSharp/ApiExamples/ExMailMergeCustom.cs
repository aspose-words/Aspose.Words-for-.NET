// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
    public class ExMailMergeCustom : ApiExampleBase
    {
        //ExStart
        //ExFor:IMailMergeDataSource
        //ExFor:IMailMergeDataSource.TableName
        //ExFor:IMailMergeDataSource.MoveNext
        //ExFor:IMailMergeDataSource.GetValue
        //ExFor:IMailMergeDataSource.GetChildDataSource
        //ExFor:MailMerge.Execute(IMailMergeDataSourceCore)
        //ExSummary:Performs mail merge from a custom data source.
        [Test] //ExSkip
        public void CustomDataSource()
        {
            // Create a destination document for the mail merge
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(" MERGEFIELD FullName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");

            // Create some data that we will use in the mail merge
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            // To be able to mail merge from your own data source, it must be wrapped
            // into an object that implements the IMailMergeDataSource interface
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // Now you can pass your data source into Aspose.Words
            doc.MailMerge.Execute(customersDataSource);

            doc.Save(ArtifactsDir + "MailMergeCustom.CustomDataSource.docx");
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
            }

            public string FullName { get; set; }
            public string Address { get; set; }
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

            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }

            private bool IsEof
            {
                get { return (mRecordIndex >= mCustomers.Count); }
            }

            private readonly CustomerList mCustomers;
            private int mRecordIndex;
        }
        //ExEnd

        //ExStart
        //ExFor:IMailMergeDataSourceRoot
        //ExFor:IMailMergeDataSourceRoot.GetDataSource(String)
        //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSourceRoot)
        //ExSummary:Performs mail merge from a custom data source with master-detail data.
        [Test] //ExSkip
        public void CustomDataSourceRoot()
        {
            // Create a document with two mail merge regions named "Washington" and "Seattle"
            Document doc = CreateSourceDocumentWithMailMergeRegions(new string[] { "Washington", "Seattle" });

            // Create two data sources
            EmployeeList employeesWashingtonBranch = new EmployeeList();
            employeesWashingtonBranch.Add(new Employee("John Doe", "Sales"));
            employeesWashingtonBranch.Add(new Employee("Jane Doe", "Management"));

            EmployeeList employeesSeattleBranch = new EmployeeList();
            employeesSeattleBranch.Add(new Employee("John Cardholder", "Management"));
            employeesSeattleBranch.Add(new Employee("Joe Bloggs", "Sales"));

            // Register our data sources by name in a data source root
            DataSourceRoot sourceRoot = new DataSourceRoot();
            sourceRoot.RegisterSource("Washington", new EmployeeListMailMergeSource(employeesWashingtonBranch));
            sourceRoot.RegisterSource("Seattle", new EmployeeListMailMergeSource(employeesSeattleBranch));

            // Since we have consecutive mail merge regions, we would normally have to perform two mail merges
            // However, one mail merge source data root call every relevant data source and merge automatically 
            doc.MailMerge.ExecuteWithRegions(sourceRoot);

            doc.Save(ArtifactsDir + "MailMergeCustom.CustomDataSourceRoot.docx");
        }

        /// <summary>
        /// Create document that contains consecutive mail merge regions, with names designated by the input array,
        /// for a data table of employees.
        /// </summary>
        private static Document CreateSourceDocumentWithMailMergeRegions(string[] regions)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            foreach (string s in regions)
            {
                builder.Writeln("\n" + s + " branch: ");
                builder.InsertField(" MERGEFIELD TableStart:" + s);
                builder.InsertField(" MERGEFIELD FullName");
                builder.Write(", ");
                builder.InsertField(" MERGEFIELD Department");
                builder.InsertField(" MERGEFIELD TableEnd:" + s);
            }

            return doc;
        }

        /// <summary>
        /// An example of a "data entity" class in your application.
        /// </summary>
        private class Employee
        {
            public Employee(string aFullName, string aDepartment)
            {
                FullName = aFullName;
                Department = aDepartment;
            }

            public string FullName { get; }
            public string Department { get; }
        }

        /// <summary>
        /// An example of a typed collection that contains your "data" objects.
        /// </summary>
        private class EmployeeList : ArrayList
        {
            public new Employee this[int index]
            {
                get { return (Employee)base[index]; }
                set { base[index] = value; }
            }
        }

        /// <summary>
        /// Data source root that can be passed directly into a mail merge which can register and contain many child data sources.
        /// These sources must all implement IMailMergeDataSource, and are registered and differentiated by a name
        /// which corresponds to a mail merge region that will read the respective data.
        /// </summary>
        private class DataSourceRoot : IMailMergeDataSourceRoot
        {
            public IMailMergeDataSource GetDataSource(string tableName)
            {
                EmployeeListMailMergeSource source = mSources[tableName];
                source.Reset();
                return mSources[tableName];
            }

            public void RegisterSource(string sourceName, EmployeeListMailMergeSource source)
            {
                mSources.Add(sourceName, source);
            }

            private readonly Dictionary<string, EmployeeListMailMergeSource> mSources = new Dictionary<string, EmployeeListMailMergeSource>();
        }

        /// <summary>
        /// Custom mail merge data source.
        /// </summary>
        private class EmployeeListMailMergeSource : IMailMergeDataSource
        {
            public EmployeeListMailMergeSource(EmployeeList employees)
            {
                mEmployees = employees;
                mRecordIndex = -1;
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

            private bool IsEof
            {
                get { return (mRecordIndex >= mEmployees.Count); }
            }

            public void Reset()
            {
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName
            {
                get { return "Employees"; }
            }

            /// <summary>
            /// Aspose.Words calls this method to get a value for every data field.
            /// </summary>
            public bool GetValue(string fieldName, out object fieldValue)
            {
                switch (fieldName)
                {
                    case "FullName":
                        fieldValue = mEmployees[mRecordIndex].FullName;
                        return true;
                    case "Department":
                        fieldValue = mEmployees[mRecordIndex].Department;
                        return true;
                    default:
                        // A field with this name was not found, 
                        // return false to the Aspose.Words mail merge engine
                        fieldValue = null;
                        return false;
                }
            }

            /// <summary>
            /// Child data sources are for nested mail merges.
            /// </summary>
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                throw new System.NotImplementedException();
            }

            private readonly EmployeeList mEmployees;
            private int mRecordIndex;
        }
        //ExEnd
    }
}