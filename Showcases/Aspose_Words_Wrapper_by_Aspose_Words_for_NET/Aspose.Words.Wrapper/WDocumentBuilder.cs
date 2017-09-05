using Aspose.Words.Fields;

namespace Aspose.Words.Wrapper
{
    public class WDocumentBuilder : DocumentBuilder
    {
        public void WriteNewLine()
        {
            Writeln();
        }

        public void WriteLine(string text)
        {
            Writeln(text);
        }

        /// <summary>
        /// Insert field by field code.
        /// </summary>
        /// <param name="fieldCode"></param>
        /// <returns></returns>
        public Field InsertFieldByFieldCode(string fieldCode)
        {
            return InsertField(fieldCode);
        }

        /// <summary>
        /// Insert field by field code and field value.
        /// </summary>
        /// <param name="fieldCode"></param>
        /// <param name="fieldValue"></param>
        /// <returns></returns>
        public Field InsertFieldWithValue(string fieldCode, string fieldValue)
        {
            return InsertField(fieldCode, fieldValue);
        }
        
        /// <summary>
        /// Insert field by field type and optionally update inserted field.
        /// </summary>
        /// <param name="fieldType"></param>
        /// <param name="updateField"></param>
        /// <returns></returns>
        public Field InsertFieldByFieldType(FieldType fieldType, bool updateField)
        {
            return InsertField(fieldType, updateField);
        }
    }
}
