using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ArtOfTest.WebAii.Core;
using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Design;
using ArtOfTest.WebAii.Design.Execution;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;
using System.Drawing;

namespace TS_CustomTestController
{
    public class TS_CustomException : Exception
    {

        public TS_CustomException(string message, string testName, string testStep, string testDescription, string errorSummary) : base(message)
        {

            GeneralUtility.WriteLog(GlobalConfig.LogFile, message);
            TS_CustomTestController.GeneralUtility.WriteLog(GlobalConfig.TestFolder + "\\TestRun_"+TS_CustomTestController.GeneralUtility.dateTimeNow+".csv", testStep + "," + testDescription + "," + "Failed" + "," + "=HYPERLINK(\"" + GlobalConfig.TestResultsFolder + "\\ErrorDetails_" + TS_CustomTestController.GeneralUtility.dateTimeNow + ".html" + "\")");

            //Create an image with the error
            Bitmap imgToWrite = Manager.Current.ActiveBrowser.Window.GetBitmap();

            GeneralUtility.WriteImage(imgToWrite, GlobalConfig.TestResultsFolder + "\\Error_Image_" + TS_CustomTestController.GeneralUtility.dateTimeNow + ".bmp");

            //create HTML file with current date time appended to filename and store in 
            GeneralUtility.WriteErrorToHtml(GlobalConfig.TestResultsFolder + "\\ErrorDetails_" + TS_CustomTestController.GeneralUtility.dateTimeNow + ".html", testName, testStep, message, "Error_Image_" + TS_CustomTestController.GeneralUtility.dateTimeNow + ".bmp", errorSummary);

            GeneralUtility.GenerateHtmlResult(GlobalConfig.TestFolder + "\\TestRun_"+TS_CustomTestController.GeneralUtility.dateTimeNow+".csv", GlobalConfig.TestFolder + "\\Results_"+TS_CustomTestController.GeneralUtility.dateTimeNow+".html", testName);
            GeneralUtility.GenerateHtmlSummaryResult(GlobalConfig.TestFolder + "\\TestSummary_"+TS_CustomTestController.GeneralUtility.dateTimeNow+".html", testName, false, GlobalConfig.TestFolder + "\\Results_"+TS_CustomTestController.GeneralUtility.dateTimeNow+".html");
        }
    }
}
