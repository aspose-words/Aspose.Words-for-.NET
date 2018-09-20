using System;

namespace ApiExamples.TestData.TestClasses
{
    public class ContractTestClass
    {
        public ManagerTestClass Manager { get; set; }
        public ClientTestClass Client { get; set; }
        public float Price { get; set; }
        public DateTime Date { get; set; }
    }
}