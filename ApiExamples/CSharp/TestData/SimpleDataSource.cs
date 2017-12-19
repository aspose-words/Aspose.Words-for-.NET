namespace ApiExamples.TestData
{
    public class SimpleDataSource
    {
        public SimpleDataSource(string name, string message)
        {
            Name = name;
            Message = message;
        }

        public string Name { get; set; }
        public string Message { get; set; }
    }
}