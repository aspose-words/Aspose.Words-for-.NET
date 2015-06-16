﻿//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Fields;
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

namespace CSharp.Mail_Merge
{
    class MailMergeFormFields
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); ;

            // Load the template document.
            Document doc = new Document(dataDir + "Template.doc");

            // Setup mail merge event handler to do the custom work.
            doc.MailMerge.FieldMergingCallback = new HandleMergeField();

            // This is the data for mail merge.
            String[] fieldNames = new String[] {"RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
                "Subject", "Body", "Urgent", "ForReview", "PleaseComment"};
            Object[] fieldValues = new Object[] {"Josh", "Jenny", "123456789", "", "Hello",
                "Test message 1", true, false, true};

            // Execute the mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the finished document.
            doc.Save(dataDir + "Template Out.doc");

            Console.WriteLine("\nMail merge performed with form fields successfully.\nFile saved at " + dataDir + "Template Out.doc");
        }

        private class HandleMergeField : IFieldMergingCallback
        {
            /// <summary>
            /// This handler is called for every mail merge field found in the document,
            ///  for every record found in the data source.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (mBuilder == null)
                    mBuilder = new DocumentBuilder(e.Document);

                // We decided that we want all boolean values to be output as check box form fields.
                if (e.FieldValue is bool)
                {
                    // Move the "cursor" to the current merge field.
                    mBuilder.MoveToMergeField(e.FieldName);

                    // It is nice to give names to check boxes. Lets generate a name such as MyField21 or so.
                    string checkBoxName = string.Format("{0}{1}", e.FieldName, e.RecordIndex);

                    // Insert a check box.
                    mBuilder.InsertCheckBox(checkBoxName, (bool)e.FieldValue, 0);

                    // Nothing else to do for this field.
                    return;
                }

                // Another example, we want the Subject field to come out as text input form field.
                if (e.FieldName == "Subject")
                {
                    mBuilder.MoveToMergeField(e.FieldName);
                    string textInputName = string.Format("{0}{1}", e.FieldName, e.RecordIndex);
                    mBuilder.InsertTextInput(textInputName, TextFormFieldType.Regular, "", (string)e.FieldValue, 0);
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }

            private DocumentBuilder mBuilder;
        }
    }
}
