using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithTextboxes
    {
        [Test]
        public void CreateALink()
        {
            //ExStart:CreateALink
            Document doc = new Document();

            Shape shape1 = new Shape(doc, ShapeType.TextBox);
            Shape shape2 = new Shape(doc, ShapeType.TextBox);

            TextBox textBox1 = shape1.TextBox;
            TextBox textBox2 = shape2.TextBox;

            if (textBox1.IsValidLinkTarget(textBox2))
                textBox1.Next = textBox2;
            //ExEnd:CreateALink
        }

        [Test]
        public void CheckSequence()
        {
            //ExStart:CheckSequence
            Document doc = new Document();

            Shape shape = new Shape(doc, ShapeType.TextBox);
            TextBox textBox = shape.TextBox;

            if (textBox.Next != null && textBox.Previous == null)
            {
                Console.WriteLine("The head of the sequence");
            }

            if (textBox.Next != null && textBox.Previous != null)
            {
                Console.WriteLine("The Middle of the sequence.");
            }

            if (textBox.Next == null && textBox.Previous != null)
            {
                Console.WriteLine("The Tail of the sequence.");
            }
            //ExEnd:CheckSequence
        }

        [Test]
        public void BreakALink()
        {
            //ExStart:BreakALink
            Document doc = new Document();

            Shape shape = new Shape(doc, ShapeType.TextBox);
            TextBox textBox = shape.TextBox;

            // Break a forward link.
            textBox.BreakForwardLink();

            // Break a forward link by setting a null.
            textBox.Next = null;

            // Break a link, which leads to this textbox.
            textBox.Previous?.BreakForwardLink();
            //ExEnd:BreakALink
        }
    }
}