using System;

namespace ApiExamples.TestData
{
    public class SignPerson
    {
        public Guid PersonId { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public byte[] Image { get; set; }
    }
}
