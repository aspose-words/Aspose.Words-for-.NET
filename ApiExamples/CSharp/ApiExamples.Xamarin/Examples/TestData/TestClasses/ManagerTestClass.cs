using System.Collections.Generic;

namespace ApiExamples.TestData.TestClasses
{
    public class ManagerTestClass
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public IEnumerable<ContractTestClass> Contracts { get; set; }
    }
}