using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CustomTestController;
using ArtOfTest.WebAii.Core;
using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;
using System.Drawing;
using ArtOfTest.WebAii.Exceptions;
using EnvDTE;
using System.Xml.Linq;
using System.Xml;
using System.Windows;


namespace TestRunner
{
    class TestRunner
    {
        static void Main(string[] args)
        {
            //kill open excel processes on stop debug mode
            //EnvDTE.DebuggerEvents debugEvents = 

            List<TestCase> TestCases = new List<TestCase>();
            string configName = "";

            new TestRunner().SetupEnvironment();

            if (args.Length == 0)
            {
                //if no arguments set a default config name
                configName = "TSLIVE1";
            }
            else
            {
                configName = args[0];
            
                Console.WriteLine("Executing Config Option -"+configName);
            }

            new TestRunner().SetConfig(TestCases,configName);


            ControllerUtility.StartTest("IE");
            ControllerUtility.WaitDelay(2000);
            new TestRunner().RunTests(TestCases);
            ControllerUtility.WaitDelay(2000);
            ControllerUtility.StopTest();

            return;
        }

        public void RunTests(List<TestCase> tsTestCases)
        {
            int startstep=0;

            foreach (TestCase tsTest in tsTestCases)
            {
                    CustomTestController.GeneralUtility.dateTimeNow = DateTime.Now.ToString("ddMMyyyy_HHmm");

                    //check if first test step is a Set Table command - if yes set flag to true and set set generic 
                    //datacount to this value
                    GlobalConfig.testCaseName = tsTest.TestCaseName;

                    if (tsTest.TestSteps[0].Action.Equals("Set Table"))
                    {
                        ExecuteSteps(tsTest, 0);
                        GlobalConfig.DataCounter = tsTest.LastRowUsed;
                        startstep = 2;
                    }
                    else
                        GlobalConfig.GlobalTCVar.LastRowUsed = tsTest.LastRowUsed;

                    if (tsTest.TestRowList != null | GlobalConfig.GlobalTCVar.LastRowUsed != 0)
                    {
                        for (int rowCtr = startstep; rowCtr <= (GlobalConfig.GlobalTCVar.LastRowUsed); rowCtr++)
                        {                            
                            ExecuteSteps(tsTest,rowCtr);
                            CustomTestController.GeneralUtility.dateTimeNow = DateTime.Now.ToString("ddMMyyyy_HHmm");
                        }
                    }
                    else
                    {
                        ExecuteSteps(tsTest);
                    }
                    GlobalConfig.GlobalTCVar = new TestCase();
            }
                
        }

        public void ExecuteSteps(TestCase TC, int currentRow=0)
        {
            bool exceptionFlag = false;

            foreach (TestStep tsStep in TC.TestSteps)
            {
                try
                {
                    Manager.Current.ActiveBrowser.RefreshDomTree();
                    Manager.Current.ActiveBrowser.WaitUntilReady();

                    //check if tsStep.Target is data-driven, if not run this step normally
                    if (tsStep.Target.Contains("["))
                    {
                        if (tsStep.Action.Equals("PassArray"))
                        {
                           CustomTestController.ControllerUtility.GenericDictionary = CustomTestController.ControllerUtility.ReplaceDictValues(TC, tsStep.Target, currentRow);                           
                        }
                        else
                        {
                            string TargetText = tsStep.Target.Substring(1, tsStep.Target.Length - 2);

                            //search for TargetText in the current row context
                            //int Colctr = TC.HeaderInfo.IndexOf(TargetText);
                            int rowctr;
                            if (currentRow==0)
                                rowctr=currentRow;
                            else
                                rowctr=currentRow-1;

                            //TS_CustomTestController.ControllerUtility.GenericDictionary = TS_CustomTestController.ControllerUtility.GenericListDictionary[rowctr];

                            CustomTestController.ControllerUtility.RunStep(tsStep.Action, TC.TestRowList[rowctr].DataValue[TargetText], tsStep.FindLogic, tsStep.Description, tsStep.Value, tsStep.BreakPoint, tsStep.StepNumber.ToString(), TC.TestCaseName);
                        }
                    }
                    else
                    {
                        //see if this step we need to call another testcase 
                        if (tsStep.Action.Equals("Execute"))
                        {
                            TestCase TestCaseStep = InitializeTestCase(GlobalConfig.TestFolder+tsStep.Target);

                            foreach (TestStep step in TestCaseStep.TestSteps)
                            {
                                Manager.Current.ActiveBrowser.RefreshDomTree();
                                Manager.Current.ActiveBrowser.WaitUntilReady();

                                if (step.Target.Contains("["))
                                {
                                    if (step.Action.Equals("PassArray"))
                                    {
                                        CustomTestController.ControllerUtility.GenericDictionary = CustomTestController.ControllerUtility.ReplaceDictValues(TestCaseStep, step.Target, currentRow);
                                    }
                                    else
                                    {
                                        string TargetText = step.Target.Substring(1, step.Target.Length - 2);

                                        //search for TargetText in the current row context
                                        if (step.Action.Equals("Verify"))
                                            TargetText = step.Target.Substring(step.Target.IndexOf('[') + 1, step.Target.Length - step.Target.IndexOf('[')-2);

                                        int Colctr = TestCaseStep.HeaderInfo.IndexOf(TargetText);

                                        CustomTestController.ControllerUtility.RunStep(step.Action, TestCaseStep.TestRowList[currentRow].DataValue[TargetText], step.FindLogic, step.Description, step.Value, step.BreakPoint, step.StepNumber.ToString(), TestCaseStep.TestCaseName);
                                    }
                                }
                                else
                                {
                                    CustomTestController.ControllerUtility.RunStep(step.Action, step.Target, step.FindLogic, step.Description, step.Value, step.BreakPoint, step.StepNumber.ToString(), TestCaseStep.TestCaseName);
                                }
                            }
                        }
                        else
                        {
                            if (tsStep.Action.Equals("Set Table") && currentRow.Equals(0))
                            {
                                CustomTestController.ControllerUtility.RunStep(tsStep.Action, tsStep.Target, tsStep.FindLogic, tsStep.Description, tsStep.Value, tsStep.BreakPoint, tsStep.StepNumber.ToString(), TC.TestCaseName);
                                TC.TestRowList = GlobalConfig.GlobalTCVar.TestRowList;
                            }
                            else
                            {
                                CustomTestController.ControllerUtility.RunStep(tsStep.Action, tsStep.Target, tsStep.FindLogic, tsStep.Description, tsStep.Value, tsStep.BreakPoint, tsStep.StepNumber.ToString(), TC.TestCaseName);
                            }
                        }
                    }
                    CustomTestController.GeneralUtility.WriteLog(GlobalConfig.LogFile, tsStep.StepNumber + "," + tsStep.Description + "," + "Passed");
                    CustomTestController.GeneralUtility.WriteLog(GlobalConfig.ResultsFolder + "\\TestRun_" + GlobalConfig.testCaseName+"_"+CustomTestController.GeneralUtility.dateTimeNow + ".csv", tsStep.StepNumber + "," + tsStep.Description + "," + "Passed");
                }
                catch (FindException Ex)
                {
                    //Trigger an error if not found (Find logic failed)
                }
                catch (ExecuteCommandException Ex)
                {
                    //Trigger an exception if Execution failed (control was found but step execution failed)
                }
                catch (Exception Ex)
                {
                    GeneralUtility.WriteLog(GlobalConfig.LogFile, Ex.Message + System.Environment.NewLine + "From Test Case Source: " + System.Environment.NewLine + GlobalConfig.testCaseName);
                    CustomTestController.GeneralUtility.WriteLog(GlobalConfig.ResultsFolder + "\\TestRun_"+GlobalConfig.testCaseName+"_" + CustomTestController.GeneralUtility.dateTimeNow + ".csv", tsStep.StepNumber.ToString() + "," + tsStep.Description + "," + "Failed" + "," + "=HYPERLINK(\"" + GlobalConfig.TestResultsFolder + "\\ErrorDetails_" + GlobalConfig.testCaseName+"_"+CustomTestController.GeneralUtility.dateTimeNow + ".html" + "\")");

                    //Create an image with the error - Save it as a jpeg!                    
                    //Bitmap imgToWrite = Manager.Current.ActiveBrowser.Window.GetBitmap();
                    System.Drawing.Image imgToWrite = (System.Drawing.Image)Manager.Current.ActiveBrowser.Window.GetBitmap();
                    GeneralUtility.WriteImage(imgToWrite, GlobalConfig.ImageResultsFolder + "\\Error_Image_"+ GlobalConfig.testCaseName+"_"+ CustomTestController.GeneralUtility.dateTimeNow + ".jpeg");

                    //create HTML file with current date time appended to filename and store in 
                    GeneralUtility.WriteErrorToHtml(GlobalConfig.TestResultsFolder + "\\ErrorDetails_"+GlobalConfig.testCaseName+"_" + CustomTestController.GeneralUtility.dateTimeNow + ".html", GlobalConfig.testCaseName, GlobalConfig.testCaseNumberError, Ex.StackTrace, GlobalConfig.ImageResultsFolder+"\\Error_Image_"+GlobalConfig.testCaseName+"_" + CustomTestController.GeneralUtility.dateTimeNow + ".jpeg", Ex.Message + System.Environment.NewLine + "From Test Case Source: " + System.Environment.NewLine + GlobalConfig.testCaseName);

                    GeneralUtility.GenerateHtmlResult(GlobalConfig.ResultsFolder + "\\TestRun_"+GlobalConfig.testCaseName+"_" + CustomTestController.GeneralUtility.dateTimeNow + ".csv", GlobalConfig.TestResultsFolder + "\\Results_"+GlobalConfig.testCaseName+"_" + CustomTestController.GeneralUtility.dateTimeNow + ".html", TC.TestCaseName);
                    GeneralUtility.GenerateHtmlSummaryResult(GlobalConfig.TestResultsFolder + "\\TestSummary_" +GlobalConfig.testCaseName+"_"+ CustomTestController.GeneralUtility.dateTimeNow + ".html", TC.TestCaseName, false, GlobalConfig.TestResultsFolder + "\\Results_" + GlobalConfig.testCaseName+"_"+CustomTestController.GeneralUtility.dateTimeNow + ".html");
                    exceptionFlag = true;
                    break;
                }
            }
            if (!exceptionFlag)
            {
                GeneralUtility.GenerateHtmlResult(GlobalConfig.ResultsFolder + "\\TestRun_" +GlobalConfig.testCaseName+"_"+ CustomTestController.GeneralUtility.dateTimeNow + ".csv", GlobalConfig.TestResultsFolder + "\\Results_" +GlobalConfig.testCaseName+"_"+ CustomTestController.GeneralUtility.dateTimeNow + ".html", TC.TestCaseName);
                GeneralUtility.GenerateHtmlSummaryResult(GlobalConfig.TestResultsFolder + "\\TestSummary_" +GlobalConfig.testCaseName+"_"+ CustomTestController.GeneralUtility.dateTimeNow + ".html", TC.TestCaseName, true, GlobalConfig.TestResultsFolder + "\\Results_" + GlobalConfig.testCaseName+"_"+CustomTestController.GeneralUtility.dateTimeNow + ".html");
            }
        }


        public void SetConfig(List<TestCase> testCases,string ConfigName)
        {   
            //Initialize global settings
           
            string dataSourcePath = GlobalConfig.SystemFolder;
            Microsoft.Office.Interop.Excel.Application excelApp;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            
            string ConfigFile;
            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Workbook workbook;

            ConfigFile = dataSourcePath + "\\TestConfig.xml";

            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(ConfigFile);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);
            
            //setup the global variables - to do: search for correct config, from command line switch or global variable
            int lastRowUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            int ctr;
            for (ctr = 0; ctr < lastRowUsed; ctr++)
            {
                if (((string)worksheet.Cells[ctr + 1, 2].Value) == null)
                {
                    break;
                }
            }

            if (ConfigName.Equals("TSLIVE1"))
            {
                GlobalConfig.SetGlobalConfig((string)worksheet.Cells[2, 1].Value, (string)worksheet.Cells[2, 2].Value, (string)worksheet.Cells[2, 4].Value, (string)worksheet.Cells[2, 5].Value, (string)worksheet.Cells[2, 3].Value, (string)worksheet.Cells[2, 6].Value, (string)worksheet.Cells[2, 8].Value);
            }
            else
            {
                for (int rowctr = 1; rowctr < lastRowUsed; rowctr++)
                {
                    if (((string)worksheet.Cells[rowctr + 1, 1].Value).Equals(ConfigName))
                    {
                        GlobalConfig.SetGlobalConfig((string)worksheet.Cells[rowctr+1, 1].Value, (string)worksheet.Cells[rowctr+1, 2].Value, (string)worksheet.Cells[rowctr+1, 4].Value, (string)worksheet.Cells[rowctr+1, 5].Value, (string)worksheet.Cells[rowctr+1, 3].Value, (string)worksheet.Cells[rowctr+1, 6].Value, (string)worksheet.Cells[rowctr+1, 8].Value);
                    }
                }
            }

            workbook.Close(true, misValue, misValue);
            excelApp.Quit();

            CustomTestController.ControllerUtility.ReleaseExcelObject(worksheet);
            CustomTestController.ControllerUtility.ReleaseExcelObject(workbook);
            //TS_CustomTestController.ControllerUtility.ReleaseExcelObject(excelApp);

            //now that we know the testdriver file name, open it and load the testcase to tsTestCase
            workbook = excelApp.Workbooks.Open(GlobalConfig.TestDriver);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);

            //Initialize test cases -
            lastRowUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            
            for (ctr = 0; ctr < lastRowUsed; ctr++)
            {
                if (((string)worksheet.Cells[ctr + 1, 2].Value) == null)
                {
                    break;
                }
            }

            TestCase tsTestCase; int testCaseCtr=0;
            for (int rowctr = 1; rowctr < lastRowUsed; rowctr++)
            {
                tsTestCase = InitializeTestCase((string)worksheet.Cells[rowctr+1, 1].Value);
                testCaseCtr++;
                testCases.Add(tsTestCase);
            }

            CustomTestController.ControllerUtility.ReleaseExcelObject(worksheet);
            CustomTestController.ControllerUtility.ReleaseExcelObject(workbook);
            CustomTestController.ControllerUtility.ReleaseExcelObject(excelApp);

            CustomTestController.GlobalConfig.ObjectRepxmldoc = XDocument.Load(GlobalConfig.ObjectRepositoryPath);
            GlobalConfig.testCaseCount = testCaseCtr;
        }

        public TestCase InitializeTestCase(string TestCaseFile)
        {
            object misValue = System.Reflection.Missing.Value;
            TestCase testCase;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            TestStep testStep;
            List<string> header=new List<string>();

            //Initialize test steps for testcase
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(TestCaseFile);
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);
           
            //return the name of the testcase from the TestCaseFile name
            char[] inputArray = TestCaseFile.ToCharArray();
            Array.Reverse(inputArray);
            string reversedName = new string(inputArray);
            
            //now search for the /
            int startPos = reversedName.IndexOf("\\");
            int endPos = reversedName.IndexOf(".")+1;

            string TestCaseName = TestCaseFile.Substring(TestCaseFile.Length - startPos, startPos-endPos);

            int lastRowUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            int ctr;
            for (ctr=0; ctr<lastRowUsed; ctr++)
            {
                if (((string)worksheet.Cells[ctr+1,2].Value)==null)
                {
                    break;
                }
            }
            lastRowUsed = ctr;
            //setup the global variables - to do: search for correct config, from command line switch or global variable
            testCase = new TestCase(TestCaseName);
            testCase.TestSteps = new List<TestStep>();

            for (int rowctr = 1; rowctr < lastRowUsed; rowctr++)
            {
                testStep = new TestStep();
 
                testStep.StepNumber = (int)worksheet.Cells[rowctr + 1, 1].Value;
                testStep.Action = (string)worksheet.Cells[rowctr + 1, 2].Value;
                testStep.Target = (string)worksheet.Cells[rowctr + 1, 3].Value;
                testStep.Description = (string)worksheet.Cells[rowctr + 1, 4].Value;
                testStep.Value = (String.IsNullOrEmpty((string)worksheet.Cells[rowctr+1,5].Value))? 0 : (int)worksheet.Cells[rowctr + 1, 5].Value;
                testStep.Value = 0;
                testStep.FindLogic = (string)worksheet.Cells[rowctr + 1, 6].Value;

                //get breakpoint value, if blank, set to false, if "yes" set to true
                if (string.IsNullOrEmpty((string)worksheet.Cells[rowctr + 1, 7].Value))
                    testStep.BreakPoint = false;
                else
                    testStep.BreakPoint = true;

                testCase.TestSteps.Add(testStep);
            }

            //Initialize datasheet for this testcase
            workbook = excelApp.Workbooks.Open(TestCaseFile);

            //point to the second worksheet
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(2);

            //return the name of the testcase from the TestCaseFile name
            inputArray = TestCaseFile.ToCharArray();
            Array.Reverse(inputArray);
            reversedName = new string(inputArray);

            //now search for the /
            startPos = reversedName.IndexOf("\\");
            endPos = reversedName.IndexOf(".") + 1;

            string Tempvar = "";
            Tempvar = (string)worksheet.Cells[1, 1].Value;
            //check if just the first column is not blank
            if (((string)worksheet.Cells[1, 1].Value) != null)
            {
                int lastColumnUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                lastRowUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                if (lastRowUsed < 1)
                {
                    //no data on 2nd worksheet
                    testCase.TestRowList = null;
                }
                else
                {
                    //setup the test data dictionary
                    testCase.TestRowList = new List<TestRowData>();

                    //get the data header                
                    for (int colheader = 1; colheader <= lastColumnUsed; colheader++)
                    {
                        header.Add((string)worksheet.Cells[1, colheader].Value);
                    }

                    for (int rowctr = 1; rowctr < lastRowUsed; rowctr++)
                    {
                        TestRowData testRowData = new TestRowData(); string ColumnHeader, ColumnValue;
                        for (int colctr = 0; colctr < lastColumnUsed; colctr++)
                        {
                            ColumnHeader = header.ElementAt<string>(colctr);
                            ColumnValue = (string)worksheet.Cells[rowctr + 1, colctr + 1].Value;

                            testRowData.DataValue.Add(ColumnHeader, ColumnValue);
                        }
                        testCase.TestRowList.Add(testRowData);
                    }

                }

                testCase.HeaderInfo = header;
                testCase.LastRowUsed = lastRowUsed - 1;
            }
            else
            {
                testCase.TestRowList = null;
                testCase.HeaderInfo = null;
                testCase.LastRowUsed = 0;
            }

            workbook.Close(true, misValue, misValue);
            excelApp.Quit();

            CustomTestController.ControllerUtility.ReleaseExcelObject(worksheet);
            CustomTestController.ControllerUtility.ReleaseExcelObject(workbook);
            CustomTestController.ControllerUtility.ReleaseExcelObject(excelApp);
            //TS_CustomTestController.GlobalConfig.ObjectRepositoryXml
            
            return testCase;
        }

        public void SetupEnvironment(string ConfigName="")
        {
            string solutionDirectory = ((EnvDTE.DTE)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.11.0")).Solution.FullName;
            solutionDirectory = System.IO.Path.GetDirectoryName(solutionDirectory);

            //always should default to current solution folder
            GlobalConfig.SolutionsDirectory = solutionDirectory;
            GlobalConfig.SystemFolder=solutionDirectory+"\\System";
            GlobalConfig.ResultsFolder = solutionDirectory + "\\ResultsFolder";
        }

        public void BreakHandler()
        {
        }
    }
}
