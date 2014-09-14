using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Web;
using System.IO;
using System.Web.UI;
using System.Windows.Forms;

namespace CustomTestController
{
    public static class GeneralUtility
    {
        public static string dateTimeNow;
        public static void WriteLog(string LogFile, string LogEntry)
        {
            System.IO.StreamWriter objWriter;           
            objWriter = new System.IO.StreamWriter(LogFile, true);

            objWriter.WriteLine(LogEntry);
            objWriter.Close();
        }

        public static void WriteErrorToHtml(string HtmlName, string TestID, string TestStep, string ErrorStack, string ScreenshotFile, string errorSummary)
        {
            System.IO.FileStream fs = new FileStream(HtmlName, FileMode.Create);

            StringWriter stringWriter = new StringWriter();
            using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
            {
                w.WriteLine("<html>");
                w.Write("<h1>");
                w.Write("TestName:"+TestID);
                w.WriteLine("</h1>");
                w.Write("<h2>");
                w.Write("Step Number:"+TestStep);
                w.WriteLine("</h2>");
                w.WriteLine("<body>");
                w.WriteLine("<img src='" + ScreenshotFile + "' style='padding:1px;border:thin solid black;' width=600 height=600  alt='" + ScreenshotFile + "' />");
                w.WriteLine("</body>");
                w.WriteLine("<br/>");
                w.WriteLine("Error: "+errorSummary);
                w.WriteLine("<br/>");
                w.WriteLine("Error stack trace:");
                w.WriteLine("<br/>");
                w.WriteLine(ErrorStack);
                w.WriteLine("</body>");
                w.WriteLine("</html>");
            }
            stringWriter.Close();
            fs.Close();
            return;
        }

        public static void WriteImage(System.Drawing.Image imgToSave, string FileNameAndPath)
        {
            imgToSave.Save(FileNameAndPath, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        public static void Generic_CopyPasteText(ArtOfTest.WebAii.Core.Manager FrameworkManager)
        {
            FrameworkManager.Desktop.KeyBoard.KeyDown(Keys.Control);
            FrameworkManager.Desktop.KeyBoard.KeyPress(Keys.A);
            FrameworkManager.Desktop.KeyBoard.KeyUp(Keys.Control);
            System.Threading.Thread.Sleep(100);

            FrameworkManager.Desktop.KeyBoard.KeyDown(Keys.Control);
            FrameworkManager.Desktop.KeyBoard.KeyPress(Keys.C);
            FrameworkManager.Desktop.KeyBoard.KeyUp(Keys.Control);
            System.Threading.Thread.Sleep(100);

            FrameworkManager.Desktop.KeyBoard.KeyDown(Keys.Control);
            FrameworkManager.Desktop.KeyBoard.KeyPress(Keys.A);
            FrameworkManager.Desktop.KeyBoard.KeyUp(Keys.Control);
            System.Threading.Thread.Sleep(100);

            FrameworkManager.Desktop.KeyBoard.KeyPress(Keys.Delete);
            System.Threading.Thread.Sleep(100);

            FrameworkManager.Desktop.KeyBoard.KeyDown(Keys.Control);
            FrameworkManager.Desktop.KeyBoard.KeyPress(Keys.V);
            FrameworkManager.Desktop.KeyBoard.KeyUp(Keys.Control);
            System.Threading.Thread.Sleep(100);
        }

        public static void GenerateHtmlResult(string inputFile, string OutputHtml, string TestID)
        {
            StreamReader sr = new StreamReader(inputFile);

            //Assemble the html header
            System.IO.FileStream fs = new FileStream(OutputHtml, FileMode.Create);

            StringWriter stringWriter = new StringWriter();
            using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
            {
                w.WriteLine("<html>");
                w.WriteLine("<img src="+GlobalConfig.SystemFolder+"\\TaskStreamLogo.gif width=150 height=50>");

                w.WriteLine("<h1>");
                w.Write("TaskStream Test Automation - Test Case Results");
                w.WriteLine("<h1>");
                w.Write("<h2>");
                w.Write("TestName:" + TestID);
                w.WriteLine("</h2>");

                w.WriteLine("<table border=1>");
                w.WriteLine("<th>Step Number</th>");
                w.WriteLine("<th>Description</th>");
                w.WriteLine("<th>Pass/Failed</th>");
                w.WriteLine("<th>Error Details</th>");
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    var values = line.Split(',');
                    w.WriteLine("<tr>");
                    foreach (string stringVal in values)
                    {
                        //1-stepnumber, 2-Description, 3-Pass/Fail, 4-error html
                        if (stringVal.Contains("html"))
                        {
                            char[] inputArray = stringVal.ToCharArray();
                            Array.Reverse(inputArray);
                            string reversedName = new string(inputArray);

                            //now search for the \
                            int startPos = reversedName.IndexOf("\\");
                            int endPos = reversedName.IndexOf(".") + 1;

                            w.WriteLine("<td> <a href=\"" + stringVal.Substring(12, stringVal.Length - 14) + "\">Error Link</a></td>");
                            //w.WriteLine("<td> <img src='" + ImageName + "' style='padding:1px;border:thin solid black;' width=600 height=600  alt='" + ImageName + "' /> </td>");
                        }
                        else
                        {
                            switch (stringVal)
                            {
                               case "Passed": w.WriteLine("<td style=\"background-color:LimeGreen \">" + stringVal + "</td>"); break;
                               case "Failed": w.WriteLine("<td style=\"background-color:Red \">" + stringVal + "</td>"); break;
                               default: w.WriteLine("<td>" + stringVal + "</td>"); break;
                            }
                        }


                    }
                    w.WriteLine("</tr>");
                }
                w.WriteLine("</table>");
                w.WriteLine("</body>");            
                w.WriteLine("</html>");
            }
            stringWriter.Close();
            fs.Close();
            return;
        }

        public static void GenerateHtmlSummaryResult(string OutputHtml, string TestID, bool Status, string TestResult)
        {
            //Assemble the html header
            System.IO.FileStream fs = new FileStream(OutputHtml, FileMode.Create);

            StringWriter stringWriter = new StringWriter();
            using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
            {
                w.WriteLine("<html>");
                w.WriteLine("<img src="+GlobalConfig.SystemFolder+"\\TaskStreamLogo.gif width=150 height=50>");

                w.WriteLine("<h1>");
                w.Write("TaskStream Test Automation - Test Summary Results");
                w.WriteLine("<h1>");
                w.Write("<h2>");
                w.Write("Execution Time:" + DateTime.Now.ToString());
                w.WriteLine("</h2>");

                w.WriteLine("<table border=1>");
                w.WriteLine("<th>Test Case Name</th>");
                w.WriteLine("<th>Pass/Failed</th>");
                w.WriteLine("<th>Results</th>");

                w.WriteLine("<tr>");
                w.WriteLine("<td>"+TestID+"</td>");
                if (Status)
                {
                    w.WriteLine("<td style=\"background-color:LimeGreen \"> Passed</td>");
                }
                else
                {
                    w.WriteLine("<td style=\"background-color:Red \"> Failed</td>"); 
                }
                w.WriteLine("<td> <a href=\"" + TestResult + "\">Result</a></td>");
                w.WriteLine("<tr>");
                w.WriteLine("</table>");
                w.WriteLine("</body>");
                w.WriteLine("</html>");
            }
            stringWriter.Close();
            fs.Close();
            return;
        }

    }
}
