using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using ArtOfTest.WebAii.Core;
using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;
using System.Reflection;
using System.Diagnostics;
using System.Xml.Linq;


namespace CustomTestController
{
    public static class ControllerUtility
    {
        public static Browser FrameElement = null;
        public static string PassValue = "";
        public static Dictionary<string, string> GenericDictionary;
        public static List<Dictionary<string, string>> GenericListDictionary;
      
        public static void NavigateToURL(string URL)
        {
            Manager.Current.ActiveBrowser.NavigateTo(URL);            
        }

        public static void ClickObject(HtmlControl objectToClick, bool clickToVisible=false) 
        {
            Manager.Current.ActiveBrowser.RefreshDomTree();
            Manager.Current.ActiveBrowser.WaitUntilReady();
            if (clickToVisible)
                objectToClick.ScrollToVisible();
            objectToClick.MouseClick(); 
        }

        public static void WaitDelay(int waitCount) 
        {
            System.Threading.Thread.Sleep(waitCount);            
        }

        public static void TypeText(string TextToType) 
        {
            Manager.Current.ActiveBrowser.RefreshDomTree();
            Manager.Current.Desktop.KeyBoard.KeyDown(Keys.Control);
            Manager.Current.Desktop.KeyBoard.KeyPress(Keys.A);
            Manager.Current.Desktop.KeyBoard.KeyUp(Keys.Control);
            System.Threading.Thread.Sleep(100);

            Manager.Current.Desktop.KeyBoard.KeyPress(Keys.Delete);
            System.Threading.Thread.Sleep(100);
            Manager.Current.Desktop.KeyBoard.TypeText(TextToType, 5, 2);            
        }

        public static Browser FindFrame(string frameName)
        {
            Manager.Current.ActiveBrowser.RefreshDomTree();
            Browser testFrame = Manager.Current.ActiveBrowser.Frames[frameName.TrimEnd()];
            return testFrame;
        }

        public static HtmlControl FindElement(string[] FindExpression, Browser frameHandle = null)
        {
            Manager.Current.ActiveBrowser.RefreshDomTree();
            Manager.Current.ActiveBrowser.WaitUntilReady();
            HtmlFindExpression ExprLogic = new HtmlFindExpression(FindExpression);
            if (frameHandle == null)
            {
                return Manager.Current.ActiveBrowser.Find.ByExpression(ExprLogic, true).As<HtmlControl>();
            }
            else
            {
                return FrameElement.Find.ByExpression(ExprLogic, true).As<HtmlControl>();
            }
        }
        
        public static HtmlSelect FindSelectElement(string[] FindExpression, Browser frameHandle = null)
        {
            Manager.Current.ActiveBrowser.RefreshDomTree();
            HtmlFindExpression ExprLogic = new HtmlFindExpression(FindExpression);
            if (frameHandle == null)
            {
                return Manager.Current.ActiveBrowser.Find.ByExpression(ExprLogic, true).As<HtmlSelect>();
            }
            else
            {
                return FrameElement.Find.ByExpression(ExprLogic, true).As<HtmlSelect>();
            }
        }

        public static HtmlSelect FindProperty(string[] FindExpression, Browser frameHandle = null)
        {
            Manager.Current.ActiveBrowser.RefreshDomTree();
            HtmlFindExpression ExprLogic = new HtmlFindExpression(FindExpression);
            if (frameHandle == null)
            {
                return Manager.Current.ActiveBrowser.Find.ByExpression(ExprLogic, true).As<HtmlSelect>();
            }
            else
            {
                HtmlSelect temp;
                temp = FrameElement.Find.ByExpression(ExprLogic, true).As<HtmlSelect>();
            
                return FrameElement.Find.ByExpression(ExprLogic, true).As<HtmlSelect>();
            }
        }

        public static HtmlDiv FindDivElement(string[] FindExpression, Browser frameHandle = null)
        {
            Manager.Current.ActiveBrowser.RefreshDomTree();
            HtmlFindExpression ExprLogic = new HtmlFindExpression(FindExpression);
            if (frameHandle == null)
            {
                return Manager.Current.ActiveBrowser.Find.ByExpression(ExprLogic, true).As<HtmlDiv>();
            }
            else
            {
                return FrameElement.Find.ByExpression(ExprLogic, true).As<HtmlDiv>();
            }
        }

        public static void StartTest(string TestBrowser)
        {
            BrowserType browserType=new BrowserType();
            switch (TestBrowser)
            {
                case "IE": browserType = BrowserType.InternetExplorer; break;
                case "FF": browserType = BrowserType.FireFox; break;
                case "C": browserType = BrowserType.Chrome; break;
            }

            //Manager newManager = new Manager(new Settings(browserType, @"c:\log\"));
            //Manager newManager = new Manager(new Settings(browserType, @"c:\log\"));
            var browserSettings = new Settings();
            browserSettings.Web.DefaultBrowser = browserType;
            Manager newManager = new Manager(browserSettings);

            newManager.Start();
            newManager.LaunchNewBrowser();
        }

        public static void StopTest()
        {
            Manager.Current.Dispose();
            GC.Collect();
        }

        public static void RunStep(string Action, string Target, string FindLogic, string Description, int Value, bool BreakPoint, string StepNumber, string TestCaseName)
        {
            bool errorFlag=false;

            //Built in delay
            WaitDelay(500);
            try
            {
                HtmlControl TargetElement; HtmlSelect SelectControl;
                if (BreakPoint)
                    Debugger.Break();

                switch (Action.TrimEnd())
                {
                    case "Navigate": NavigateToURL(Target); break;
                    case "Type": TargetElement = FindElement(findRepositoryElement(FindLogic).Split(','), FrameElement);
                        ClickObject(TargetElement);
                        TypeText(Target); break;
                    case "Click": TargetElement = FindElement(findRepositoryElement(FindLogic).Split(','), FrameElement);
                        ClickObject(TargetElement);
                        break;
                    case "ClickToVisible": TargetElement = FindElement(findRepositoryElement(FindLogic).Split(','), FrameElement);
                        ClickObject(TargetElement,true);
                        break;
                    case "Set Frame": FrameElement=Target.TrimEnd().Equals("null") ? null : FindFrame(Target); break;
                    case "Call": RunCodedStep(Target); break;
                    case "PassValue": CustomTestController.ControllerUtility.PassValue = Target; break;
                    case "Select": SelectControl = FindSelectElement(findRepositoryElement(FindLogic).Split(','), FrameElement);
                        SelectControl.SelectByText(Target); break;
                    case "ConnectToBrowser": Manager.Current.SetNewBrowserTracking(true);
                                             Manager.Current.WaitForNewBrowserConnect(Target, true, 5000);
                                             Manager.Current.ActiveBrowser.WaitUntilReady(); break;
                    case "Verify": TargetElement = FindElement(findRepositoryElement(FindLogic).Split(','), FrameElement);
                        errorFlag = PropertyVerify(Target, TargetElement.BaseElement);
                        if (!errorFlag) throw new Exception("Verify Failed -" + Target);
                        break;
                    case "Wait": WaitDelay(Value); break;
                    case "Set Table": SetTable(GlobalConfig.TestFolder+Target); break;
                    default: errorFlag = true; throw new Exception("Incorrect Keyword -"+Target);
                }
            }
            catch (Exception e)
            {
                string ErrorString = "";
                GlobalConfig.testCaseNumberError = StepNumber;
                GlobalConfig.testCaseName = TestCaseName;
                if (!errorFlag)
                {
                    ErrorString = String.IsNullOrEmpty(Action) ? ErrorString += " (Action=NULL)" : " (Action=" + Action + ")";
                    ErrorString = String.IsNullOrEmpty(Target) ? ErrorString += " (Target=NULL)" : ErrorString+=" (Target=" + Target + ")";
                    ErrorString = String.IsNullOrEmpty(FindLogic) ? ErrorString += " (FindLogic=NULL)" : ErrorString+=" (FindID=" + FindLogic + ", " + "XML Object Repository return = " + findRepositoryElement(FindLogic) +  ")";
                    throw new Exception("Error: Could not run -" + ErrorString+ " Error:"+e.Message);
                }
                else throw;
            }

        }

        public static void ReleaseExcelObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void RunCodedStep(string ClassAndMethod)
        {
                //Hard code this for now - 
                //Assembly assembly = Assembly.LoadFile("C:\\Work In Progress - Projects\\TS_CustomTestController\\TS_CustomTestController\\bin\\Debug\\TS_CustomTestController.dll");
                Assembly assembly = Assembly.LoadFile(GlobalConfig.SolutionsDirectory+"\\TS_CustomTestController\\bin\\Debug\\TS_CustomTestController.dll");
                int ClassNameDelimiter = ClassAndMethod.IndexOf(".");
                int strLength = ClassAndMethod.Length;
                int classNameLength = (ClassAndMethod.Length - ClassNameDelimiter );
                int MethodLength = strLength - (ClassNameDelimiter +1);
                
                string ClassName = ClassAndMethod.Substring(0, ClassNameDelimiter);                
                string MethodName = ClassAndMethod.Substring(ClassNameDelimiter + 1, MethodLength);

                Type CodedStepClassType = Type.GetType("TS_CustomTestController."+ClassName);
                
                if (CodedStepClassType != null)
                {
                    MethodInfo methodInfo = CodedStepClassType.GetMethod(MethodName);
                    if (methodInfo != null)
                    {
                        try
                        {
                            object result = null;
                            ParameterInfo[] parameters = methodInfo.GetParameters();
                            object classInstance = Activator.CreateInstance(CodedStepClassType, null);
                            if (parameters.Length == 0)
                            {
                                result = methodInfo.Invoke(classInstance, null);
                            }
                            else
                            {
                                object[] parametersArray = new object[] { "Temp" };

                                result = methodInfo.Invoke(classInstance, parametersArray);
                            }
                        }
                        catch (TargetInvocationException ex)
                        {
                            throw new Exception("Class or Method - " + ClassAndMethod + " does not exist or threw an exception.");
                        }
                    }
                    else
                    {
                        throw new Exception("Class or Method - " + ClassAndMethod + " does not exist or threw an exception.");
                    }
                }
                else
                {
                    throw new Exception("Class or Method - " + ClassAndMethod + " does not exist or threw an exception.");
                }
        }

        public static Dictionary<string, string> ReplaceDictValues(TestCase TC, string PassValues, int CurrentRow)
        {
            Dictionary<string, string> DictValues = new Dictionary<string, string>();
            string Key, Value;
            foreach (string Target in PassValues.Split(','))
            {
                if (Target.Contains('['))
                    //Colctr = TC.HeaderInfo.IndexOf(Target.Substring(1,Target.Length-2));
                    Key = Target.Substring(1, Target.Length - 2);
                else
                    //Colctr = TC.HeaderInfo.IndexOf(Target.Substring(0, Target.Length));
                    Key = Target.Substring(0, Target.Length);

                //Key = Target.Substring(1, Target.Length - 2);
                Value = TC.TestRowList[CurrentRow].DataValue[Key];
                DictValues.Add(Key, Value);
            }
            return DictValues;
        }

        public static bool PropertyVerify(string verifyString, Element targetElement)
        {
            //break the string apart from the = symbol, left hand side is the property, right hand is value
            //Example Target - innertext="Lesson Builder"
            string propertyVar = verifyString.Substring(0, verifyString.IndexOf('='));
            string valueVar = verifyString.Substring(verifyString.IndexOf('=') + 1,verifyString.Length-(verifyString.IndexOf('=')+1));

            if (targetElement.GetAttributeValueOrEmpty(propertyVar).Equals(valueVar))
                return true;
            else
                return false;
        }

        //Search XML document by elementID and return the text in curresponding "findlogic" child node
        public static string findRepositoryElement(string elementID)
        {

            string returnFindLogic;
            string ErrorString;
        
            var elementInfo = from e in CustomTestController.GlobalConfig.ObjectRepxmldoc.Descendants("UIElement")
                                where (string)e.Attribute("ID").Value == elementID
                                select (string)e.Element("findlogic");              

            if (elementInfo.Count() > 1)
            {
                ErrorString = "Too many arguments found in the repository with that ID";
                throw new Exception (ErrorString );
            }

            if (elementInfo.Count() == 0)
            {
                ErrorString = "No elements found in the repository with that ID";
                throw new Exception(ErrorString);

            }

            returnFindLogic = elementInfo.FirstOrDefault().ToString();

            return returnFindLogic;                         
        }

        public static void SetTable(string TestDataFile)
        {
            object misValue = System.Reflection.Missing.Value;
            TestCase testCase;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            List<string> header = new List<string>();
            List<TestRowData> TestRowList=new List<TestRowData>();

            //Initialize test steps for testcase
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(TestDataFile);
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);

            //return the name of the testcase from the TestCaseFile name
            char[] inputArray = TestDataFile.ToCharArray();
            Array.Reverse(inputArray);
            string reversedName = new string(inputArray);

            //now search for the /
            int startPos = reversedName.IndexOf("\\");
            int endPos = reversedName.IndexOf(".") + 1;

            string TestCaseName = TestDataFile.Substring(TestDataFile.Length - startPos, startPos - endPos);

            int lastRowUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            int ctr;
            for (ctr = 0; ctr < lastRowUsed; ctr++)
            {
                if (((string)worksheet.Cells[ctr + 1, 2].Value) == null)
                {
                    break;
                }
            }
            lastRowUsed = ctr;
            testCase = new TestCase("DataSheet");
            string columnHeader = "";
            //check if just the first column is not blank
            if (((string)worksheet.Cells[1, 1].Value) != null)
            {
                int lastColumnUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                lastRowUsed = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

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
                         columnHeader = columnHeader + ","+ColumnHeader;
                     }
                     testCase.TestRowList.Add(testRowData);
                  }
                columnHeader = columnHeader.Substring(1, columnHeader.Length - 1);
            }
            testCase.HeaderInfo = header;
            testCase.LastRowUsed = lastRowUsed - 1;

            GlobalConfig.GlobalTCVar = testCase;

            CustomTestController.ControllerUtility.GenericDictionary = null;
            Dictionary<string, string> DictValues;
            GenericListDictionary = new List<Dictionary<string, string>>();
            
            for (int rowctr = 0; rowctr < lastRowUsed-1; rowctr++)
            {
                DictValues =  new Dictionary<string, string>();
                string Key, Value;
                foreach (string Target in testCase.HeaderInfo)
                {
                    if (Target.Contains('['))
                        Key = Target.Substring(1, Target.Length - 2);
                    else
                        Key = Target.Substring(0, Target.Length);

                    Value = testCase.TestRowList[rowctr].DataValue[Key];
                    DictValues.Add(Key, Value);
                }
                GenericListDictionary.Add(DictValues);
            }
            
            workbook.Close(true, misValue, misValue);
            excelApp.Quit();

            CustomTestController.ControllerUtility.ReleaseExcelObject(worksheet);
            CustomTestController.ControllerUtility.ReleaseExcelObject(workbook);
            CustomTestController.ControllerUtility.ReleaseExcelObject(excelApp);
            //TS_CustomTestController.GlobalConfig.ObjectRepositoryXml

            return;

        }
    }
}
