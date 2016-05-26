// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;

using Aspose.Words;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExVariableCollection : ApiExampleBase
    {
        [Test] 
        public void AddEx()
        {
            //ExStart
            //ExFor:VariableCollection.Add
            //ExSummary:Shows how to create document variables and add them to a document's variable collection.
            Document doc = new Document(MyDir + "Document.doc");

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
            Document doc = new Document(MyDir + "Document.doc");

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
            Document doc = new Document(MyDir + "Document.doc");

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
            Document doc = new Document(MyDir + "Document.doc");

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

        [Test]
        public void IndexOfKeyEx()
        {
            //ExStart
            //ExFor:VariableCollection.IndexOfKey
            //ExSummary:Shows how to get the index of a key.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Variables.Add("doc", "Word processing document");
            doc.Variables.Add("docx", "Word processing document");
            doc.Variables.Add("txt", "Plain text file");
            doc.Variables.Add("bmp", "Image");
            doc.Variables.Add("png", "Image");

            Console.WriteLine(doc.Variables.IndexOfKey("bmp")); // 0
            Console.WriteLine(doc.Variables.IndexOfKey("txt")); // 4
            //ExEnd
        }

        [Test]
        public void RemoveEx()
        {
            //ExStart
            //ExFor:VariableCollection.Remove
            //ExSummary:Shows how to remove an element from a document's variable collection by key.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Variables.Add("doc", "Word processing document");
            doc.Variables.Add("docx", "Word processing document");
            doc.Variables.Add("txt", "Plain text file");
            doc.Variables.Add("bmp", "Image");
            doc.Variables.Add("png", "Image");

            doc.Variables.Remove("bmp");
            Console.WriteLine(doc.Variables.Count); // 4
            //ExEnd
        }

        [Test]
        public void RemoveAtEx()
        {
            //ExStart
            //ExFor:VariableCollection.RemoveAt
            //ExSummary:Shows how to remove an element from a document's variable collection by index.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Variables.Add("doc", "Word processing document");
            doc.Variables.Add("docx", "Word processing document");
            doc.Variables.Add("txt", "Plain text file");
            doc.Variables.Add("bmp", "Image");
            doc.Variables.Add("png", "Image");

            int index = doc.Variables.IndexOfKey("bmp");
            doc.Variables.RemoveAt(index);
            Console.WriteLine(doc.Variables.Count); // 4
            //ExEnd
        }
    }
}
