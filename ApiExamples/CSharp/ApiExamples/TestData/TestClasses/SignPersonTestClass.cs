using System;

namespace ApiExamples.TestData.TestClasses
{
    public class SignPersonTestClass
    {
        public Guid PersonId { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public byte[] Image { get; set; }

        public SignPersonTestClass(Guid guid, string name, string position, byte[] image)
        {
            this.PersonId = guid;
            this.Name = name;
            this.Position = position;
            this.Image = image;
        }
    }
}
