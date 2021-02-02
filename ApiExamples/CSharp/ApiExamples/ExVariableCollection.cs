// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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

            // Every document has a collection of key/value pair variables, which we can add items to.
            variables.Add("Home address", "123 Main St.");
            variables.Add("City", "London");
            variables.Add("Bedrooms", "3");

            Assert.AreEqual(3, variables.Count);

            // We can display the values of variables in the document body using DOCVARIABLE fields.
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldDocVariable field = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            field.VariableName = "Home address";
            field.Update();

            Assert.AreEqual("123 Main St.", field.Result);

            // Assigning values to existing keys will update them.
            variables.Add("Home address", "456 Queen St.");

            // We will then have to update DOCVARIABLE fields to ensure they display an up-to-date value.
            Assert.AreEqual("123 Main St.", field.Result);

            field.Update();

            Assert.AreEqual("456 Queen St.", field.Result);

            // Verify that the document variables with a certain name or value exist.
            Assert.True(variables.Contains("City"));
            Assert.True(variables.Any(v => v.Value == "London"));

            // The collection of variables automatically sorts variables alphabetically by name.
            Assert.AreEqual(0, variables.IndexOfKey("Bedrooms"));
            Assert.AreEqual(1, variables.IndexOfKey("City"));
            Assert.AreEqual(2, variables.IndexOfKey("Home address"));

            // Enumerate over the collection of variables.
            using (IEnumerator<KeyValuePair<string, string>> enumerator = doc.Variables.GetEnumerator())
                while (enumerator.MoveNext())
                    Console.WriteLine($"Name: {enumerator.Current.Key}, Value: {enumerator.Current.Value}");

            // Below are three ways of removing document variables from a collection.
            // 1 -  By name:
            variables.Remove("City");

            Assert.False(variables.Contains("City"));

            // 2 -  By index:
            variables.RemoveAt(1);

            Assert.False(variables.Contains("Home address"));

            // 3 -  Clear the whole collection at once:
            variables.Clear();

            Assert.That(variables, Is.Empty);
            //ExEnd
        }
    }
}