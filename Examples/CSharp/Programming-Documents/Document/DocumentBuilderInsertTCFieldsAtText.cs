using System.IO;
using Aspose.Words;
using System;
using System.Drawing;
using Aspose.Words.Tables;
using System.Text.RegularExpressions;
using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertTCFieldsAtText
    {
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertTCFieldsAtText
            Document doc = new Document();

            FindReplaceOptions options = new FindReplaceOptions();

            // Highlight newly inserted content.
            options.ApplyFont.HighlightColor = Color.DarkOrange;
            options.ReplacingCallback =  new InsertTCFieldHandler("Chapter 1", "\\l 1");

            // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
            doc.Range.Replace(new Regex("The Beginning"), "", options);
            //ExEnd:DocumentBuilderInsertTCFieldsAtText
          
        }     
    }
    //ExStart:InsertTCFieldHandler
    public class InsertTCFieldHandler : IReplacingCallback
    {
        // Store the text and switches to be used for the TC fields.
        private string mFieldText;
        private string mFieldSwitches;

        /// <summary>
        /// The switches to use for each TC field. Can be an empty string or null.
        /// </summary>
        public InsertTCFieldHandler(string switches)
            : this(string.Empty, switches)
        {
            mFieldSwitches = switches;
        }

        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty string or null.
        /// </summary>
        public InsertTCFieldHandler(string text, string switches)
        {
            mFieldText = text;
            mFieldSwitches = switches;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Create a builder to insert the field.
            DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
            // Move to the first node of the match.
            builder.MoveTo(args.MatchNode);

            // If the user specified text to be used in the field as display text then use that, otherwise use the 
            // match string as the display text.
            string insertText;

            if (!string.IsNullOrEmpty(mFieldText))
                insertText = mFieldText;
            else
                insertText = args.Match.Value;

            // Insert the TC field before this node using the specified string as the display text and user defined switches.
            builder.InsertField(string.Format("TC \"{0}\" {1}", insertText, mFieldSwitches));

            // We have done what we want so skip replacement.
            return ReplaceAction.Skip;
        }
    }
    //ExEnd:InsertTCFieldHandler
}
