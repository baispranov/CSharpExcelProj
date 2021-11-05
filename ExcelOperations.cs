using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelProject
{
    class ExcelOperations
    {
        public bool ApplyDataValidation()
        {
            try
            {
                // Step 1 : Open Excel 
                string excelFileName = Environment.CurrentDirectory +@"\Input.xlsx";
                Console.WriteLine();
                // Excel application Interface 
                Excel.Application IApplication;
                // Excel Workbook interface
                Excel.Workbook IWorkbook;
                // Excel worsheet interface 
                Excel.Worksheet IWorksheet;
          

                IApplication = new Excel.Application();
                Object missing = System.Reflection.Missing.Value;
                IApplication.Visible = true;
                IApplication.WindowState = Excel.XlWindowState.xlMaximized;
                IWorkbook = IApplication.Workbooks.Open(excelFileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                IWorksheet = (Worksheet)IWorkbook.Worksheets.get_Item(1);

                // Activate Sheet1 in alrady opened Excel File - Activate sheet 2 
                IWorksheet = (Worksheet)IWorkbook.Sheets["Sheet1"];
                IWorksheet.Activate();

                // Step 2 : Apply Data validation to A1 Cell - Entire column in Excel Sheet 
                Range cell = IWorksheet.Rows.Cells[1, 1];
                cell.Value = "Data Validation Field.";
                cell.EntireColumn.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, Excel.XlFormatConditionOperator.xlBetween, @"1,2,3,4,5");
                
                // Step 3 : Save Excel Sheet 
                IWorkbook.Save();
                IWorkbook.Close();
                
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("Exception message is "+ ex.Message );
            }
        }
    }
}
