using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Core;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;


namespace CustomTestController
{
    public class ListObject_Find
    {
        public HtmlControl Get_Item_Link(string IName) //Returns the lesson's name link, searches based on lesson name
        {

            //HtmlControl lessonLink = Pages.Taskstream_LAT_Default.FrameTSmain.Find.ByExpression<HtmlControl>("innertext=" + IName, "tagname=a");//Pages.Taskstream_LAT_Default.FrameTSmain.Find.ByExpression<HtmlControl>("innertext=" + lessonNameX,"tagname=a");
            HtmlControl lessonLink = CustomTestController.ControllerUtility.FindElement(("innertext=" + IName+",tagname=a").Split(','), CustomTestController.ControllerUtility.FrameElement);
            //if (lessonLink == null)
            //{
            //    lessonLink = Pages.Taskstream_LAT_Default.FrameTSBottom.Find.ByExpression<HtmlControl>("innertext=" + IName, "tagname=a");
            //}

            return lessonLink;

        }

        public void Click_Item_link(string IName)
        {
            Get_Item_Link(IName).Click();
        }

        public HtmlControl Get_Items_Contatiner(string IName)//Returns the parent LI or TR tag (container) based on name
        {


            HtmlControl aLink = Get_Item_Link(IName); //get lesson link

            HtmlControl LI = aLink.Parent<HtmlListItem>();//Find parent container of the lesson link
          
            if(LI == null)
                LI = aLink.Parent<HtmlTableRow>();
            
            return LI;

        }

        public HtmlInputControl Get_DeleteButton(string lName)//Returns the 'delete' button of a lesson, based on lesson name
        {
            HtmlControl LI = Get_Items_Contatiner(lName);//gives me access to the Li tag 
           
            HtmlInputControl deleteButton = LI.Find.ByAttributes<HtmlInputControl>("value=Delete"); //Withing the correct LI tag get me the button
            
            return deleteButton;
                           

        }

        public void Click_Item_DeleteButton(string IName)
        {
            Get_DeleteButton(IName).Click();
        }

    }

}
