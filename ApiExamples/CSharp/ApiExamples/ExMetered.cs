// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExMetered : ApiExampleBase
    {
        [Test]
        public void TestMeteredUsage()
        {
            Assert.Throws<InvalidOperationException>(MeteredUsage);
        }

        public void MeteredUsage()
        {
            //ExStart
            //ExFor:Metered
            //ExFor:Metered.#ctor
            //ExFor:Metered.GetConsumptionCredit
            //ExFor:Metered.GetConsumptionQuantity
            //ExFor:Metered.SetMeteredKey(String, String)
            //ExSummary:Shows how to activate a Metered license and track credit/consumption.
            // Set a public and private key for a new Metered instance
            Metered metered = new Metered();
            metered.SetMeteredKey("MyPublicKey", "MyPrivateKey");
            
            // Print credit/usage 
            Console.WriteLine($"Credit before operation: {Metered.GetConsumptionCredit()}");
            Console.WriteLine($"Consumption quantity before operation: {Metered.GetConsumptionQuantity()}");

            // Do something
            Document doc = new Document(MyDir + "Document.doc");

            // Print credit/usage to see how much was spent
            Console.WriteLine($"Credit after operation: {Metered.GetConsumptionCredit()}");
            Console.WriteLine($"Consumption quantity after operation: {Metered.GetConsumptionQuantity()}");
            //ExEnd
        }
    }
}