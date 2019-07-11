using Aspose.Words.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Linked_Textboxes
{
    class WorkingWithLinkedTextboxes
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithShapes();
            CreateALink(dataDir);
            CheckSequence(dataDir);
            BreakALink(dataDir);
        }

        public static void CreateALink(string dataDir)
        {
            // ExStart:CreateALink
            Document doc = new Document();
            Shape shape1 = new Shape(doc, ShapeType.TextBox);
            Shape shape2 = new Shape(doc, ShapeType.TextBox);

            TextBox textBox1 = shape1.TextBox;
            TextBox textBox2 = shape2.TextBox;

            if (textBox1.IsValidLinkTarget(textBox2))
                textBox1.Next = textBox2;
            // ExEnd:CreateALink
        }

        public static void CheckSequence(string dataDir)
        {
            // ExStart:CheckSequence
            Document doc = new Document();
            Shape shape = new Shape(doc, ShapeType.TextBox);
            TextBox textBox = shape.TextBox;

            if ((textBox.Next != null) && (textBox.Previous == null))
            {
                Console.WriteLine("The head of the sequence");
            }

            if ((textBox.Next != null) && (textBox.Previous != null))
            {
                Console.WriteLine("The Middle of the sequence.");
            }

            if ((textBox.Next == null) && (textBox.Previous != null))
            {
                Console.WriteLine("The Tail of the sequence.");
            }
            // ExEnd:CheckSequence
        }

        public static void BreakALink(string dataDir)
        {
            // ExStart:BreakALink
            Document doc = new Document();
            Shape shape = new Shape(doc, ShapeType.TextBox);
            TextBox textBox = shape.TextBox;

            // Break a forward link
            textBox.BreakForwardLink();

            // Break a forward link by setting a null
            textBox.Next = null;

            // Break a link, which leads to this textbox
            if (textBox.Previous != null)
                textBox.Previous.BreakForwardLink();
            // ExEnd:BreakALink
        }
    }
}
