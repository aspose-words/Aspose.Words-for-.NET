using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Load_Options
{
    public class WorkingWithHtmlLoadOptions : DocsExamplesBase
    {
        [Test]
        public void PreferredControlType()
        {
            //ExStart:LoadHtmlElementsWithPreferredControlType
            const string html = @"
                <html>
                    <select name='ComboBox' size='1'>
                        <option value='val1'>item1</option>
                        <option value='val2'></option>                        
                    </select>
                </html>
            ";

            HtmlLoadOptions loadOptions = new HtmlLoadOptions { PreferredControlType = HtmlControlType.StructuredDocumentTag };

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithHtmlLoadOptions.PreferredControlType.docx", SaveFormat.Docx);
            //ExEnd:LoadHtmlElementsWithPreferredControlType
        }
    }
}