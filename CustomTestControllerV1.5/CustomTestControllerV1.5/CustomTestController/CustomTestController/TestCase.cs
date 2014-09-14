using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CustomTestController
{
    public class TestCase
    {
        public List<TestStep> TestSteps;
        public string TestCaseName;
        public List<TestRowData> TestRowList;
        public List<string> HeaderInfo;
        public int LastRowUsed;

        public TestCase() : this("") { }
   
        public TestCase(string testCaseName)
        {
            TestCaseName = testCaseName;
        }
    }
}
