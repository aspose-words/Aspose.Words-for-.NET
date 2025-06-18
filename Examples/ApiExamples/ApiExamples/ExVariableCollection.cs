// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
            //ExFor:VariableCollection.Item(Int32)
            //ExFor:VariableCollection.Item(String)
            //ExSummary:Shows how to work with a document's variable collection.
            Document doc = new Document();
            VariableCollection variables = doc.Variables;

            // Every document has a collection of key/value pair variables, which we can add items to.
            variables.Add("Home address", "123 Main St.");
            variables.Add("City", "London");
            variables.Add("Bedrooms", "3");

            Assert.That(variables.Count, Is.EqualTo(3));

            // We can display the values of variables in the document body using DOCVARIABLE fields.
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldDocVariable field = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            field.VariableName = "Home address";
            field.Update();

            Assert.That(field.Result, Is.EqualTo("123 Main St."));

            // Assigning values to existing keys will update them.
            variables.Add("Home address", "456 Queen St.");

            // We will then have to update DOCVARIABLE fields to ensure they display an up-to-date value.
            Assert.That(field.Result, Is.EqualTo("123 Main St."));

            field.Update();

            Assert.That(field.Result, Is.EqualTo("456 Queen St."));

            // Verify that the document variables with a certain name or value exist.
            Assert.That(variables.Contains("City"), Is.True);
            Assert.That(variables.Any(v => v.Value == "London"), Is.True);

            // The collection of variables automatically sorts variables alphabetically by name.
            Assert.That(variables.IndexOfKey("Bedrooms"), Is.EqualTo(0));
            Assert.That(variables.IndexOfKey("City"), Is.EqualTo(1));
            Assert.That(variables.IndexOfKey("Home address"), Is.EqualTo(2));

            Assert.That(variables[0], Is.EqualTo("3"));
            Assert.That(variables["City"], Is.EqualTo("London"));

            // Enumerate over the collection of variables.
            using (IEnumerator<KeyValuePair<string, string>> enumerator = doc.Variables.GetEnumerator())
                while (enumerator.MoveNext())
                    Console.WriteLine($"Name: {enumerator.Current.Key}, Value: {enumerator.Current.Value}");

            // Below are three ways of removing document variables from a collection.
            // 1 -  By name:
            variables.Remove("City");

            Assert.That(variables.Contains("City"), Is.False);

            // 2 -  By index:
            variables.RemoveAt(1);

            Assert.That(variables.Contains("Home address"), Is.False);

            // 3 -  Clear the whole collection at once:
            variables.Clear();

            Assert.That(variables.Count, Is.EqualTo(0));
            //ExEnd
        }
    }
}