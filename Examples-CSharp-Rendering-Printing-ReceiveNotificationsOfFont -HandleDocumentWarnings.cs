// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
public class HandleDocumentWarnings : IWarningCallback
{
    /// <summary>
    /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
    /// potential issue during document procssing. The callback can be set to listen for warnings generated during document
    /// load and/or document save.
    /// </summary>
    public void Warning(WarningInfo info)
    {
        // We are only interested in fonts being substituted.
        if (info.WarningType == WarningType.FontSubstitution)
        {
            Console.WriteLine("Font substitution: " + info.Description);
        }
    }
}
