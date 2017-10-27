using AventStack.ExtentReports;
using log4net.Repository.Hierarchy;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWAUTCSharpFramework.Utilities
{
    class GenericKeywords : Common
    {
        public static IWebDriver driver;
        public static ExtentReports extent;
        public static ExtentTest parent;
        public static ExtentTest child;
        public static String identifier;
        public static String locator;
        public static String locatorDescription;
        public static String outputDirectory;
        public static String currentExcelBook;
        public static String mainWindow;
        public static String currentBrowser = "";
        public static Logger logger;
        public static int currentTestCaseNumber;
        public static int currentExcelSheet;
        public static int currentStep;
        public static int failureNo;
        public static int screenshotNo; public static int rowCount;
        public static int colCount;
       // public static Common.identifierType idType;
        public static IWebElement webElement;
        public static Boolean testFailure = false;
        public static Boolean loadFailure = false;
        public static int temp = 1;
        public static String testStatus = "";
        public static int testCaseDataRow;
        public static int textLoadWaitTime;
        public static int elementLoadWaitTime;
        public static int implicitlyWaitTime;
        public static int pageLoadWaitTime = 0;
        public static int testCaseRow;
  //      public static final String XSLT_FILE_CoverPage = ".\\xsltfiles\\CoverPage.xslt";
  //public static final String XSLT_FILE_SummaryPage = ".\\xsltfiles\\SummaryReport.xslt";
  //public static final String XSLT_FILE_PdfPage = ".\\data\\PdfReport.xslt";
  //public static final ArrayList<String> testCaseNames = new ArrayList();
  //      public static ArrayList<String> testCaseDataSets = new ArrayList();
        public static Boolean windowreadyStateStatus = true;
        public static int testSuccessCount = 0;
        public static int testFailureCount = 0;
        public static int testSkippedCount = 0;
        public static String timeStamp = "";
        public static Boolean testCaseExecutionStatus = false;
        public static Boolean webElementIsPresent = false;



        public GenericKeywords() { }



        //public static enum platFormName
        //{
        //    IOS,
        //    ANDROID;
        //}

        //public static void openApp()
        //{
        //    String deviceName = getConfigProperty("DeviceName").toString().trim();
        //    String platForm = getConfigProperty("PlatFormName").toString().trim();
        //    String platFormVersion = getConfigProperty("PlatformVersion").toString().trim();
        //    String appName = getConfigProperty("AppName").toString().trim();


        //    String ip = getConfigProperty("IpAddress").toString().trim();
        //    String portNumber = getConfigProperty("PortNumber").toString().trim();
        //    platFormName b = platFormName.valueOf(platForm.toUpperCase());


        //    writeToLogFile("INFO", "Opening " + appName + " Application...");
        //    try
        //    {
        //        DesiredCapabilities capabilities = new DesiredCapabilities();
        //        capabilities.setCapability("newCommandTimeout", getConfigProperty("AppiumTimeOut").toString().trim());
        //        switch (b)
        //        {
        //            case IOS:
        //                break;

        //            case ANDROID:
        //                capabilities.setCapability("platformName", platForm);
        //                capabilities.setCapability("platformVersion", platFormVersion);
        //                capabilities.setCapability("deviceName", deviceName);
        //                driver = new AppiumDriver(new URL("http://" + ip + ":" + portNumber + "/wd/hub"), capabilities);
        //        }




        //        elementLoadWaitTime = Integer.parseInt(getConfigProperty("ElementLoadWaitTime").toString().trim());
        //        textLoadWaitTime = Integer.parseInt(getConfigProperty("TextLoadWaitTime").toString().trim());
        //        pageLoadWaitTime = Integer.parseInt(getConfigProperty("PageLoadWaitTime").toString().trim());
        //        implicitlyWaitTime = Integer.parseInt(getConfigProperty("ImplicitlyWaitTime").toString().trim());
        //        driver.manage().timeouts().implicitlyWait(Integer.parseInt(getConfigProperty("ImplicitlyWaitTime")), TimeUnit.SECONDS);

        //        writeToLogFile("INFO", "Time out set");
        //        writeToLogFile("INFO", "Application: Open Successful: " + appName);
        //        testReporter("Green", "Open Application: ''" + appName + "''");

        //    }
        //    catch (TimeoutException e)
        //    {
        //        testStepFailed("Page fail to load within in " + getConfigProperty("pageLoadWaitTime") + " seconds");
        //    }
        //    catch (Exception e)
        //    {
        //        writeToLogFile("ERROR", "Browser: Open Failure/Navigation cancelled, please check the application window.");
        //        writeToLogFile("Error", e.toString());
        //        testReporter("Red", e.toString());
        //        testStepFailed("Open App : AppName");
        //    }
        //}
    }
}
