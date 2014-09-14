using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CustomTestController
{
    public class TestStep
    {
        public int StepNumber { get; set;}
        public string Action { get; set; }
        public string Target { get; set; }
        public string Description { get; set; }
        public int Value { get; set; }
        public string FindLogic { get; set; }
        public bool BreakPoint;
    }
}
