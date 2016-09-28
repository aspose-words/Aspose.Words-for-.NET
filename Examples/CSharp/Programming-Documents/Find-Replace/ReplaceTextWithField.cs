using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using System;
using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceTextWithField
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();
            string fileName = "Field.ReplaceTextWithFields.doc";

            Document doc = new Document(dataDir + fileName);

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceTextWithFieldHandler(FieldType.FieldMergeField);

            // Replace any "PlaceHolderX" instances in the document (where X is a number) with a merge field.
            doc.Range.Replace(new Regex(@"PlaceHolder(\d+)"), "", options);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            doc.Save(dataDir);

            Console.WriteLine("\nText replaced with field successfully.\nFile saved at " + dataDir);
        }
    }

    public class ReplaceTextWithFieldHandler : IReplacingCallback
    {
        public ReplaceTextWithFieldHandler(FieldType type)
        {
            mFieldType = type;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            ArrayList runs = FindAndSplitMatchRuns(args);

            // Create DocumentBuilder which is used to insert the field.
            DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
            builder.MoveTo((Run)runs[runs.Count - 1]);

            // Calculate the name of the field from the FieldType enumeration by removing the first instance of "Field" from the text. 
            // This works for almost all of the field types.
            string fieldName = mFieldType.ToString().ToUpper().Substring(5);

            // Insert the field into the document using the specified field type and the match text as the field name.
            // If the fields you are inserting do not require this extra parameter then it can be removed from the string below.
            builder.InsertField(string.Format("{0} {1}", fieldName, args.Match.Groups[0]));

            // Now remove all runs in the sequence.
            foreach (Run run in runs)
                run.Remove();

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.Skip;
        }

        /// <summary>
        /// Finds and splits the match runs and returns them in an ArrayList.
        /// </summary>
        public ArrayList FindAndSplitMatchRuns(ReplacingArgs args)
        {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = args.MatchNode;

            // The first (and may be the only) run can contain text before the match, 
            // in this case it is necessary to split the run.
            if (args.MatchOffset > 0)
                currentNode = SplitRun((Run)currentNode, args.MatchOffset);

            // This array is used to store all nodes of the match for further removing.
            ArrayList runs = new ArrayList();

            // Find all runs that contain parts of the match string.
            int remainingLength = args.Match.Value.Length;
            while (
                (remainingLength > 0) &&
                (currentNode != null) &&
                (currentNode.GetText().Length <= remainingLength))
            {
                runs.Add(currentNode);
                remainingLength = remainingLength - currentNode.GetText().Length;

                // Select the next Run node. 
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do
                {
                    currentNode = currentNode.NextSibling;
                }
                while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
            }

            // Split the last run that contains the match if there is any text left.
            if ((currentNode != null) && (remainingLength > 0))
            {
                SplitRun((Run)currentNode, remainingLength);
                runs.Add(currentNode);
            }

            return runs;
        }

        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;

        }

        private FieldType mFieldType;
    }
}
