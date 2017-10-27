using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using log4net;
using log4net.Repository.Hierarchy;
using log4net.Core;
using log4net.Appender;
using log4net.Layout;
using Excel = Microsoft.Office.Interop.Excel;
using NUnit.Framework;
namespace SWAUTCSharpFramework
{

    public class Common
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static ExtentReports extent;
        public static ExtentTest parent;
        public static ExtentTest child;
        public static int currentStep;
        public static int row;
        public static int testCaseRow;
        public static bool testFailure = false;
        public static int testCaseDataRow;
        public static List<string> testCaseDataSets = new List<string> { };
        public static List<string> testCaseNames = new List<string> { };
        public static int testCaseDataNo;
        public static int rowCount;
        public static int colCount;
        public static int testCaseexecutionNo = 0;
        public static String outputDirectory;
        public static bool testCaseExecutionStatus = false;
        public static List<String> testCases = new List<string> { };
        public static IWebDriver browser;
        public static int screenshotNo;
        public static int elementLoadWaitTime;
        public static string Reportdate;
        public static string wanted_path;

        public static IWebDriver getdriver(String browserName)
        {
            IWebDriver driver = null;
            if (browserName.ToUpper().Equals("CHROME"))
            {
                try
                {

                    wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
                    wanted_path = wanted_path.Replace("file:\\", "");
                    string pathString = System.IO.Path.Combine(wanted_path, "BrowserDrivers");

                    ChromeOptions options = new ChromeOptions();
                    options.AddArguments("--start-maximized");

                    driver = new ChromeDriver(pathString, options);
                    log.Info("Exceuting in chrome browser");
                }
                catch (Exception ex)
                {
                    log.Error("Unable to launch chrome browser");
                    Console.WriteLine(ex);
                }

            }
            if (browserName.ToUpper().Equals("IE"))
            {

                wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
                wanted_path = wanted_path.Replace("file:\\", "");
                string pathString = System.IO.Path.Combine(wanted_path, "BrowserDrivers");

                driver = new InternetExplorerDriver(pathString);
                log.Info("Exceuting in Internet Explorer browser");
                driver.Manage().Window.Maximize();
            }
            if (browserName.ToUpper().Equals("FIREFOX"))
            {
                try
                {

                    driver = new FirefoxDriver();
                    log.Info("Exceuting in Firefox browser");
                    driver.Manage().Window.Maximize();

                }
                catch (Exception ex)
                {
                    log.Error("Unable to launch firefox browser");
                    Console.WriteLine(ex);
                }
            }
            return driver;
        }

        public static void Reports()
        {
            Reportdate = System.DateTime.Now.ToString(@"yyyy-MM-dd hh:mm:ss tt");
            Reportdate = Reportdate.Replace(" ", "_");
            Reportdate = Reportdate.Replace("-", "_");
            Reportdate = Reportdate.Replace(":", "_");
            wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
            wanted_path = wanted_path.Replace("file:\\", "");
            string folderName = System.IO.Path.Combine(wanted_path, "TestResults");
            string pathString = System.IO.Path.Combine(folderName, Reportdate);
            Common.outputDirectory = pathString;
            System.IO.Directory.CreateDirectory(pathString);

            string newPathString = System.IO.Path.Combine(pathString, "ExecutionReport.html");

            //creating a log object for creating a log file each time inside the Test Result folder
            ILog _log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            string newPathString1 = System.IO.Path.Combine(pathString, "ExecutionLog.txt");

            ChangeFilePath("RollingFileAppender", newPathString1);
            log.Info("Test Output Directory creation successful " + pathString);
            log.Info("Log File creation successful " + newPathString1);
            extent = new ExtentReports(newPathString);


        }
        public static void copyFiles()
        {
            string wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
            wanted_path = wanted_path.Replace("file:\\", "");
            string folderName = System.IO.Path.Combine(wanted_path, "TestResults");
            string pathString = System.IO.Path.Combine(folderName, Reportdate);
            string jenkinsPath = System.IO.Path.Combine(wanted_path, "JenkinsResults");
            DirectoryInfo TestResults = new DirectoryInfo(pathString);
            DirectoryInfo JenkinsResults = new DirectoryInfo(jenkinsPath);
            CopyFilesRecursively(TestResults, JenkinsResults);
        }

        public static void CopyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
        {
            foreach (FileInfo fi in source.GetFiles())
            {
                //Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                foreach (FileInfo fi in diSourceSubDir.GetFiles())
                {
                    //Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                    fi.CopyTo(Path.Combine(nextTargetSubDir.ToString(), fi.Name), true);
                }
            }
        }

        //To change the log file path,by deault it writes to a specific file location

        public static void ChangeFilePath(string appenderName, string newFilename)
        {
            log4net.Repository.ILoggerRepository repository = log4net.LogManager.GetRepository();
            foreach (log4net.Appender.IAppender appender in repository.GetAppenders())
            {
                if (appender.Name.CompareTo(appenderName) == 0 && appender is log4net.Appender.FileAppender)
                {
                    log4net.Appender.FileAppender fileAppender = (log4net.Appender.FileAppender)appender;
                    fileAppender.File = newFilename;//System.IO.Path.Combine(fileAppender.File, newFilename);
                    fileAppender.ActivateOptions();
                }
            }
        }

        public static void updateTestDataSet(String testCaseName)
        {
            DataDriver1.updateTestDataSet1(testCaseName);
        }

        public static void CloseReport()
        {
            extent.EndTest(parent);
            extent.Flush();
        }

        public static int returnRowNumber(String Label)
        {

            return DataDriver1.returnRowNo(2, Label);

        }
        public static String retrieve(String Label)
        {
            return DataDriver1.retrieve(Common.testCaseRow, Common.testCaseDataRow, Label);

        }

        public static void testStepPassed(String Message)
        {
            testReporter("Green", Message);
        }

        public static void testStepInfo(String errMessage)
        {
            testReporter("blue", errMessage);

        }

        public static void testStepFailed(String errMessage)
        {
            testReporter("Red", errMessage);
            takeScreenshot("Failed");

        }

        public static void testStepInfoStart(string testDataSet)
        {

            child = extent.StartTest(testDataSet);
            // child.Log(LogStatus.Info, "########### Start of Test Case Data Set: " + testDataSet + " ###########");
            log.Info("########### Start of Test Case Data Set: " + testDataSet + " ###########");

        }


        public static void testStepInfoEnd(String testDataSet)
        {
            //child.Log(LogStatus.Info, "########### End of Test Case Data Set: " + testDataSet + " ###########");
            parent.AppendChild(child);
            currentStep = 0;
            log.Info("########### End of Test Case Data Set: " + testDataSet + " ###########");
        }


        public static void testReporter(String color, String report)
        {
            currentStep += 1;
            String ct = color.ToLower();
            switch (ct)
            {
                case "green":
                    child.Log(LogStatus.Pass, "<font color=green>" + currentStep + ". " + report + "</font><br/>"); break;
                case "blue": child.Log(LogStatus.Info, "<font color=blue>" + currentStep + ". " + report + "</font><br/>"); break;
                case "orange": child.Log(LogStatus.Warning, "<font color=orange>" + currentStep + ". " + report + "</font><br/>"); break;
                case "red": child.Log(LogStatus.Fail, "<font color=red>" + currentStep + ". " + report + "</font><br/>"); break;
                case "white":
                    child.Log(LogStatus.Info, report); break;
            }

        }

        public static void screenShot(String filename)
        {
            String strPath = Common.outputDirectory;
            string pathString = System.IO.Path.Combine(strPath, "ScreenShots");
            System.IO.Directory.CreateDirectory(pathString);


            Screenshot ss = ((ITakesScreenshot)browser).GetScreenshot();

            //Use it as you want now
            string screenshot = ss.AsBase64EncodedString;
            byte[] screenshotAsByteArray = ss.AsByteArray;
            String strScreenPath = System.IO.Path.Combine(pathString, filename + ".png");
            ss.SaveAsFile(strScreenPath, ImageFormat.Png); //use any of the built in image formating
            ss.ToString();

        }

        public static void takeScreenshot(String comment)
        {
            try
            {
                screenshotNo += 1;
                screenShot("Screenshot" + screenshotNo);
                String strPath = Common.outputDirectory;
                string pathString = System.IO.Path.Combine(strPath, "ScreenShots");
                String strScreenPath = System.IO.Path.Combine(pathString, "Screenshot" + screenshotNo + ".png");
                strScreenPath = System.IO.Path.Combine(pathString, "Screenshot" + screenshotNo + ".png");
                Common.child.Log(LogStatus.Info, "Check ScreenShot Below: " + Common.child.AddScreenCapture(strScreenPath));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
        public static List<String> TestCaseSettingExel(string filepath)
        {
            List<string> rowValue = new List<string> { };
            List<string> execute = new List<string> { };
            var ExcelFilePath = filepath; ;
            Microsoft.Office.Interop.Excel.Application xlApp = new Excel.Application();
            try
            {
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelFilePath);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;


                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    rowValue.Clear();

                    for (int j = 1; j <= colCount; j++)
                    {
                        try
                        {
                            rowValue.Add(xlRange.Cells[i, j].Value2.ToString());

                        }
                        catch (Exception ex)
                        {
                            break;
                        }
                    }


                    if (rowValue[4] == "Yes")
                    {

                        string testexec = rowValue[2].ToString();
                        execute.Add(testexec);
                    }
                }


            }
            catch (Exception e)
            {
                e.StackTrace.ToString();
            }


            return execute;


        }

        public static string GetConfigProperty(string data)
        {

            string wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
            wanted_path = wanted_path.Replace("file:\\","");
            string wanted_path1 = wanted_path.Replace("\\", "/");
            string pathString = System.IO.Path.Combine(wanted_path1, "TestData");
            string newPathString = System.IO.Path.Combine(pathString, "TestConfiguration.xls");

            newPathString = newPathString.Replace("\\", "/");

            //string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DotnetWorkspace\SWAUTCSharpFramework\SeleniumFirst\TestData\TestConfiguration.xls;" +
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+newPathString+";" +
           @"Extended Properties='Excel 12.0 Xml;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
               
                     connection.Open();
                    OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                    using (OleDbDataReader dr = command.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            var row1Col0 = dr[0];
                            string Exceldata = dr[0].ToString();
                            if (Exceldata.Equals(data))
                            {
                                return dr[1].ToString();
                            }

                            System.Threading.Thread.Sleep(500);
                        }
                    }

              

               
            }

            return null;
        }


    }
}
