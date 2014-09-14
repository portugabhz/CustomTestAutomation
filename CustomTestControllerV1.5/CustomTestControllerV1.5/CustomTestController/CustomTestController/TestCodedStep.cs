using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CustomTestController;

namespace CustomTestController
{
    public class TestCodedStep
    {
        public TestCodedStep()
        {
        }

        public void ClickPlatform()
        {
            if (CustomTestController.ControllerUtility.GenericDictionary["MultiPlatform"].Equals("Yes"))
            {
                string platform = CustomTestController.ControllerUtility.GenericDictionary["TargetPlatform"];

                CustomTestController.ControllerUtility.WaitDelay(2000);
                CustomTestController.ControllerUtility.ClickObject(CustomTestController.ControllerUtility.FindElement(("tagname=input,id=loginID").Split(',')));
                CustomTestController.ControllerUtility.WaitDelay(2000);
                CustomTestController.ControllerUtility.ClickObject(CustomTestController.ControllerUtility.FindElement(("tagname=img,alt=~Learning Achievement System").Split(',')));
            }
            else
            {
                CustomTestController.ControllerUtility.WaitDelay(2000);
                CustomTestController.ControllerUtility.ClickObject(CustomTestController.ControllerUtility.FindElement(("tagname=input,id=loginID").Split(',')));
            }
        }

    }
}
