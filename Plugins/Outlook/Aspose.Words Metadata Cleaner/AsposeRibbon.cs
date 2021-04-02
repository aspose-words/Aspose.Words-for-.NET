using Microsoft.Office.Tools.Ribbon;

namespace Aspose.Words_Metadata_Cleaner
{
    public partial class AsposeRibbon
    {
        private void AsposeRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (Globals.ThisAddIn.EnableAsposeWordsMetadata == false)
                CB_EnableAsposeWordsMetadataCleaner.Checked = false;
            else
                CB_EnableAsposeWordsMetadataCleaner.Checked = true;
        }

        private void CB_EnableAsposeWordsMetadataCleaner_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.EnableAsposeWordsMetadata = CB_EnableAsposeWordsMetadataCleaner.Checked;
        }
    }
}
