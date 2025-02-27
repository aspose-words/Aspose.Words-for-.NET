﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class ConvertToMhtmlEmail : TestUtil
    {
        [Test, Ignore("")]
        public static void ConvertToMhtmlEmailFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");

            // Save into a memory stream in MHTML format.
            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it.
            stream.Position = 0;

            // Create an Aspose.Network MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, new MhtmlLoadOptions());
            message.From = "from@gmail.com";
            message.To = "to@gmail.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.gmail.com";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.Auto;
            
            client.Send(message);
        }
    }
}
