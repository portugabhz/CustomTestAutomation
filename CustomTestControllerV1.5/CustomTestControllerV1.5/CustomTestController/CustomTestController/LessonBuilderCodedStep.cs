using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CustomTestController;

using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Core;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
//using ArtOfTest.WebAii.Design;
//using ArtOfTest.WebAii.Design.Execution;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;

namespace CustomTestController
{
    public class LessonBuilderCodedStep
    {
        public LessonBuilderCodedStep()
        {
        }

        public void ClickFoundLesson()
        {
            string LessonName = CustomTestController.ControllerUtility.GenericDictionary["Name"].ToString();
            new ListObject_Find().Click_Item_link(LessonName);
        }


        public void DeleteLesson()
        {

            string itemName = CustomTestController.ControllerUtility.GenericDictionary["Name"].ToString();

            // var deleteButton = new FindListObjects().Get_DeleteButton_InList(lessonName);

            new ListObject_Find().Click_Item_DeleteButton(itemName);
            ListObject_Delete_ClickOnOK();
            //deleteButton.Click();
        }

        public void ListObject_Delete_ClickOnOK()
        {

            // Click 'OK button'

            HtmlControl btnOK;
            //btnOK = Pages.Taskstream_LAT_Default.FrameTSmain.Find.ByExpression<HtmlControl>("TextContent=OK", "tagname=span");
            btnOK = CustomTestController.ControllerUtility.FindElement(("TextContent=OK, tagname=span").Split(','), CustomTestController.ControllerUtility.FrameElement);

            //if (btnOK == null)
            //    btnOK = Pages.Taskstream_LAT_Default.FrameTSBottom.Find.ByExpression<HtmlControl>("TextContent=OK", "tagname=span");

            btnOK.Click(false);

        }

        
        public void ListObject_Delete_ValidateAlertMessage()
        {
            //string MessageText = Data["AlertMessage"].ToString();
            string MessageText = CustomTestController.ControllerUtility.GenericDictionary["AlertMessage"].ToString();

            //HtmlDiv testControl = Pages.Taskstream_LAT_Default.FrameTSmain.Find.ByExpression<HtmlDiv>("TextContent=^" + MessageText, "tagname=div");
            HtmlDiv testControl = CustomTestController.ControllerUtility.FindDivElement(("TextContent=^" + MessageText + ",tagname=div").Split(','), CustomTestController.ControllerUtility.FrameElement);

            //if (testControl == null)
            //    testControl = Pages.Taskstream_LAT_Default.FrameTSBottom.Find.ByExpression<HtmlDiv>("TextContent=^" + MessageText, "tagname=div");

            Assert.IsNotNull(testControl);
        }        
        
    }
}
