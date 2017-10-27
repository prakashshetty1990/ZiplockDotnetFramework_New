using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace SWAUTCSharpFramework
{
    public class DataDriver1 : Common
    {
        public static Excel.Application xlApp;
        public static Excel.Workbook xlWorkbook;
        public static Excel._Worksheet xlWorksheet;
        public static Excel.Range xlRange;


        public static String getData(int row, int col)
        {
            try
            {
                return xlRange.Cells[row, col].Value2.ToString();
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public static int returnRowNumber(String Label)
        {
            return returnRowNo(2, Label);
        }

        public static int returnRowNo(int colIndex, String rowLabel)
        {
            bool flag = true;
            int temp = 0;
            while (flag)
            {
                temp++;
                if (getData(temp, colIndex).Trim().Equals(rowLabel.Trim()))
                {
                    flag = false;
                    return temp;
                }

            }
            return 1;
        }


        public static String retrieve(int intLabelRow, int intDataRow, String colLabel)
        {
            return getData(intDataRow, returnColNo(intLabelRow, colLabel));
        }

        public static int returnColNo(int datasetNo, String colLabel)
        {
            bool flag = true;
            int temp = 0;
            while (flag)
            {
                temp++;
                if (getData(datasetNo, temp).Trim().Equals(colLabel.Trim()))
                {
                    flag = false;
                    return temp;
                }
            }
            return 0;
        }



        public static void updateTestDataSet1(String testCaseName)
        {
            String testCaseDataSet = "";
            String executionFlag = null;
            bool flag = false;
            string wanted_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)));
            wanted_path = wanted_path.Replace("file:\\", "");
            string pathString = System.IO.Path.Combine(wanted_path, "TestData");
            string newPathString = System.IO.Path.Combine(pathString, "TestData.xls");
            try
            {
                List<string> rowValue = new List<string> { };
                string ExcelFilePath = newPathString;
              //  string ExcelFilePath = "C:\\DotnetWorkspace\\SeleniumFirst\\SeleniumFirst\\TestData\\TestData.xls";

                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(ExcelFilePath);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                Common.testCaseDataSets.Clear();
                for (int caseRow = 1; caseRow <= rowCount; caseRow++)
                {                    
                    string str = "";
                    try
                    {
                         str = xlRange.Cells[caseRow, 2].Value2.ToString();
                    }
                    catch (Exception ex)
                    {
                        //do nothing
                    }
                    if (testCaseName.Equals(str))
                    {

                        for (int DataRow = caseRow; DataRow <= rowCount; DataRow++)
                        {
                            Common.testCaseRow = caseRow;
                            testCaseDataSet = xlRange.Cells[DataRow, 2].Value2.ToString();
                            executionFlag = xlRange.Cells[DataRow, 3].Value2.ToString();
                            if ((testCaseDataSet.StartsWith(testCaseName)) && (executionFlag.ToUpper().Equals("YES")))
                            {
                                Common.testCaseDataSets.Add(testCaseDataSet);
                            }
                            else if (testCaseDataSet.Length == 0)
                            {
                                flag = true;
                                break;
                            }

                        }
                        if (flag)
                        {
                            break;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
      
    }


}
