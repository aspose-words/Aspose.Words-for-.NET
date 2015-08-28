// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello World!");

            doc.Save("Create word processing document.docx");
        }
    }
}
