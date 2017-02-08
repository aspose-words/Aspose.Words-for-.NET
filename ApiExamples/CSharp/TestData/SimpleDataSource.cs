namespace ApiExamples.TestData
{
    public class SimpleDataSource
    {
        public SimpleDataSource(string name, string message)
        {
            this.Name = name;
            this.Message = message;
        }

        public string Name { get; set; }

        public string Message { get; set; }
    }
}