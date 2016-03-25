using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("Test.docx");
            foreach (Paragraph par in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                par.ParagraphBreakFont.Hidden = false;
                foreach (Run run in par.GetChildNodes(NodeType.Run, true))
                {
                    if (run.Font.Hidden)
                        run.Font.Hidden = false;
                }
            }
            doc.Save("Test.docx");
        }
    }
}
