// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExVariableCollection : ApiExampleBase
    {
        [Test]
        public void Primer()
        {
            //ExStart
            //ExFor:Document.Variables
            //ExFor:VariableCollection
            //ExFor:VariableCollection.Add
            //ExFor:VariableCollection.Clear
            //ExFor:VariableCollection.Contains
            //ExFor:VariableCollection.Count
            //ExFor:VariableCollection.GetEnumerator
            //ExFor:VariableCollection.IndexOfKey
            //ExFor:VariableCollection.Remove
            //ExFor:VariableCollection.RemoveAt
            //ExSummary:Shows how to work with a document's variable collection.
            Document doc = new Document();
            VariableCollection variables = doc.Variables;

            // Documents have a variable collection to which name/value pairs can be added
            variables.Add("Home address", "123 Main St.");
            variables.Add("City", "London");
            variables.Add("Bedrooms", "3");

            Assert.AreEqual(3, variables.Count);

            // Variables can be referenced and have their values presented in the document by DOCVARIABLE fields
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldDocVariable field = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            field.VariableName = "Home address";
            field.Update();

            Assert.AreEqual("123 Main St.", field.Result);

            // Assigning values to existing keys will update them
            variables.Add("Home address", "456 Queen St.");

            // DOCVARIABLE fields also need to be updated in order to show an accurate up to date value
            field.Update();

            Assert.AreEqual("456 Queen St.", field.Result);

            // The existence of variables can be looked up either by name or value like this
            Assert.True(variables.Contains("City"));
            Assert.True(variables.Any(v => v.Value == "London"));

            // Variables are automatically sorted in alphabetical order
            Assert.AreEqual(0, variables.IndexOfKey("Bedrooms"));
            Assert.AreEqual(1, variables.IndexOfKey("City"));
            Assert.AreEqual(2, variables.IndexOfKey("Home address"));

            // Enumerate over the collection of variables
            using (IEnumerator<KeyValuePair<string, string>> enumerator = doc.Variables.GetEnumerator())
                while (enumerator.MoveNext())
                    Console.WriteLine($"Name: {enumerator.Current.Key}, Value: {enumerator.Current.Value}");

            // Variables can be removed either by name or index, or the entire collection can be cleared at once
            variables.Remove("City");

            Assert.False(variables.Contains("City"));

            variables.RemoveAt(1);

            Assert.False(variables.Contains("Home address"));

            variables.Clear();

            Assert.That(variables, Is.Empty);
            //ExEnd
        }
    }
}