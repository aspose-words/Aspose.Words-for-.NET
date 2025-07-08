// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
            //ExFor:Metered.IsMeteredLicensed
            //ExFor:Metered.GetProductName
            //ExSummary:Shows how to activate a Metered license and track credit/consumption.
            // Create a new Metered license, and then print its usage statistics.
            Metered metered = new Metered();
            metered.SetMeteredKey("MyPublicKey", "MyPrivateKey");

            Console.WriteLine(string.Format("Is metered license accepted: {0}", Metered.IsMeteredLicensed()));
            Console.WriteLine(string.Format("Product name: {0}", metered.GetProductName()));
            Console.WriteLine(string.Format("Credit before operation: {0}", Metered.GetConsumptionCredit()));
            Console.WriteLine(string.Format("Consumption quantity before operation: {0}", Metered.GetConsumptionQuantity()));

            // Operate using Aspose.Words, and then print our metered stats again to see how much we spent.
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "Metered.Usage.pdf");

            // Aspose Metered Licensing mechanism does not send the usage data to purchase server every time,
            // you need to use waiting.
            System.Threading.Thread.Sleep(10000);

            Console.WriteLine(string.Format("Credit after operation: {0}", Metered.GetConsumptionCredit()));
            Console.WriteLine(string.Format("Consumption quantity after operation: {0}", Metered.GetConsumptionQuantity()));
            //ExEnd
        }
    }
}