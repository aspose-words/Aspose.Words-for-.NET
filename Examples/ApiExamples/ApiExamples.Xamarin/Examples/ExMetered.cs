// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
            Assert.Throws<InvalidOperationException>(Usage);
        }

        public void Usage()
        {
            //ExStart
            //ExFor:Metered
            //ExFor:Metered.#ctor
            //ExFor:Metered.GetConsumptionCredit
            //ExFor:Metered.GetConsumptionQuantity
            //ExFor:Metered.SetMeteredKey(String, String)
            //ExSummary:Shows how to activate a Metered license and track credit/consumption.
            // Create a new Metered license, and then print its usage statistics.
            Metered metered = new Metered();
            metered.SetMeteredKey("MyPublicKey", "MyPrivateKey");
            
            Console.WriteLine($"Credit before operation: {Metered.GetConsumptionCredit()}");
            Console.WriteLine($"Consumption quantity before operation: {Metered.GetConsumptionQuantity()}");

            // Operate using Aspose.Words, and then print our metered stats again to see how much we spent.
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "Metered.Usage.pdf");

            Console.WriteLine($"Credit after operation: {Metered.GetConsumptionCredit()}");
            Console.WriteLine($"Consumption quantity after operation: {Metered.GetConsumptionQuantity()}");
            //ExEnd
        }
    }
}