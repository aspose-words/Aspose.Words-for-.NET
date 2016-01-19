// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.VariableCollection
{
    [TestFixture]
    public class ExVariableCollection : QaTestsBase
    {
        [Test] 
        public void AddEx()
        {
            //ExStart
            //ExFor:VariableCollection.Add
            //ExSummary:Shows how to add DictionaryEntry instances to a document's VariableCollection.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Variables.Add("doc", "Word processing document");
            doc.Variables.Add("docx", "Word processing document");
            doc.Variables.Add("txt", "Word processing document");
            // Duplicate values can be stored but adding a duplicate name overwrites the old one.
            doc.Variables.Add("txt", "Plain text file");
            doc.Variables.Add("bmp", "Image");
            doc.Variables.Add("png", "Image");
            //ExEnd
        }

        [Test]
        public void ClearEx()
        {
            //ExStart
            //ExFor:VariableCollection.Clear
            //ExSummary:Shows how to clear all document variables from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Variables.Add("doc", "Word processing document");
            doc.Variables.Add("docx", "Word processing document");
            doc.Variables.Add("txt", "Plain text file");
            doc.Variables.Add("bmp", "Image");
            doc.Variables.Add("png", "Image");

            doc.Variables.Clear();
            Console.WriteLine(doc.Variables.Count); // 0
            //ExEnd
        }

        [Test]
        public void ContainsEx()
        {
            //ExStart
            //ExFor:VariableCollection.Contains
            //ExSummary:Shows how to check if a collection of document variables contains a key.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Variables.Add("doc", "Word processing document");
           
            Console.WriteLine(doc.Variables.Contains("doc")); // True
            Console.WriteLine(doc.Variables.Contains("Word processing document")); // False
            //ExEnd
        }

        [Test]
        public void GetEnumeratorEx()
        {
            //ExStart
            //ExFor:VariableCollection.GetEnumerator
            //ExSummary:Shows how to obtain an enumerator from a collection of document variables and use it.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Variables.Add("doc", "Word processing document");
            doc.Variables.Add("docx", "Word processing document");
            doc.Variables.Add("txt", "Plain text file");
            doc.Variables.Add("bmp", "Image");
            doc.Variables.Add("png", "Image");

            var enumerator = doc.Variables.GetEnumerator();

            while (enumerator.MoveNext())
            {
                DictionaryEntry de = (DictionaryEntry)enumerator.Current;
                Console.WriteLine("Name: {0}, Value: {1}", de.Key, de.Value);
            }
            //ExEnd
        }
    }
}
