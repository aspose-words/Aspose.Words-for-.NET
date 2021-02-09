using System;

namespace DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects
{
    //ExStart:Contract
    public class Contract
    {
        public Manager Manager { get; set; }
        public Client Client { get; set; }
        public float Price { get; set; }
        public DateTime Date { get; set; }
    }
    //ExEnd:Contract 
}