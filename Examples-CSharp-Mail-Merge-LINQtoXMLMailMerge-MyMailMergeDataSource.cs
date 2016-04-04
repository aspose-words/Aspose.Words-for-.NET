// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
public class MyMailMergeDataSource : IMailMergeDataSource
{
    /// <summary>
    /// Creates a new instance of a custom mail merge data source.
    /// </summary>
    /// <param name="data">Data returned from a LINQ query.</param>
    public MyMailMergeDataSource(IEnumerable data)
    {
        mEnumerator = data.GetEnumerator();
    }
    
    /// <summary>
    /// Creates a new instance of a custom mail merge data source, for mail merge with regions.
    /// </summary>
    /// <param name="data">Data returned from a LINQ query.</param>
    /// <param name="tableName">Name of the data source is only used when you perform mail merge with regions. 
    /// If you prefer to use the simple mail merge then use constructor with one parameter.</param>
    public MyMailMergeDataSource(IEnumerable data, string tableName)
    {
        mEnumerator = data.GetEnumerator();
        mTableName = tableName;
    }
    
    /// <summary>
    /// Aspose.Words calls this method to get a value for every data field.
    /// 
    /// This is a simple "generic" implementation of a data source that can work over 
    /// any IEnumerable collection. This implementation assumes that the merge field
    /// name in the document matches the name of a public property on the object
    /// in the collection and uses reflection to get the value of the property.
    /// </summary>
    public bool GetValue(string fieldName, out object fieldValue)
    {
        // Use reflection to get the property by name from the current object.
        object obj = mEnumerator.Current;

        Type curentRecordType = obj.GetType();
        PropertyInfo property = curentRecordType.GetProperty(fieldName);
        if (property != null)
        {
            fieldValue = property.GetValue(obj, null);
            return true;
        }

        // Return False to the Aspose.Words mail merge engine to indicate the field was not found.
        fieldValue = null;
        return false;
    }
    
    /// <summary>
    /// Moves to the next record in the collection.
    /// </summary>
    public bool MoveNext()
    {
        return mEnumerator.MoveNext();
    }
    
    /// <summary>
    /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
    /// </summary>
    public string TableName
    {
        get { return mTableName; }
    }
    
    public IMailMergeDataSource GetChildDataSource(string tableName)
    {
        return null;
    }

    private readonly IEnumerator mEnumerator;
    private readonly string mTableName;
}
