namespace ApiExamples.TestData.TestClasses
{
    public class MessageTestClass
    {
        public string Name { get; set; }
        public string Message { get; set; }

        public MessageTestClass(string name, string message)
        {
            Name = name;
            Message = message;
        }
    }
}