using System.Collections.Generic;

namespace DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects
{
    //ExStart:Manager
    public class Manager
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public byte[] Photo { get; set; }
        public IEnumerable<Contract> Contracts { get; set; }
    }
    //ExEnd:Manager
}