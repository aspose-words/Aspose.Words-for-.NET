using Aspose.Words.Fields;

namespace Aspose.Words.Wrapper
{
    public class WParagraph : Paragraph
    {
        public WParagraph(DocumentBase document) : base(document)
        {
        }

        public Field AppendFieldByCode(string fieldCode) 
        {
            return AppendField(fieldCode);
        }

        public Field AppendFieldWithValue(string fieldCode, string fieldValue)
        {
            return AppendField(fieldCode, fieldValue);
        }

        public Run GetRun(int runNumber)
        {
            return runNumber >= 0 && runNumber < Runs.Count 
                ? Runs[runNumber] 
                : null;
        }
    }
}
