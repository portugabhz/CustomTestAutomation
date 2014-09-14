using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace CustomTestController
{
    public static class GlobalConfig
    {
        public static string ConfigName;
        public static string StartURL;
        public static string TestResultsFolder;
        public static string ImageResultsFolder;
        public static string TestFolder;
        public static string LogFile;
        public static string TestDriver;
        public static string SystemFolder;
        public static string ResultsFolder;
        public static string SolutionsDirectory;
        public static string ObjectRepositoryPath;
        public static XDocument ObjectRepxmldoc;
        public static XDocument ObjectRepXML;
        public static int testCaseCount;
        public static string testCaseNumberError;
        public static string testCaseName;
        public static int DataCounter;
        public static TestCase GlobalTCVar=new TestCase();

        public static void SetGlobalConfig(string strConfigName, string strStartURL, string strTestResultsFolder, string strImageResultsFolder,string strTestFolder, string strTestDriver, string strObjectRep, bool setInitialTestDriver=false)
        {
            //make this relative path - relative to SystemFolder (solutions folder)
            
            ConfigName = strConfigName;
            StartURL = strStartURL;
            TestResultsFolder = SolutionsDirectory+strTestResultsFolder;
            ImageResultsFolder = SolutionsDirectory+strImageResultsFolder;
            TestFolder = SolutionsDirectory+strTestFolder;
            LogFile = TestFolder + "\\LogFile.txt";
            ObjectRepositoryPath = TestFolder + "\\ObjectRepository" + strObjectRep;
            if (!setInitialTestDriver)
                TestDriver = TestFolder + strTestDriver;
            else
                TestDriver = TestFolder + strTestDriver;
        }
    }
}
