using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;


namespace TestAutomationPromoter
{
    public struct ExcelObject
    {
        public string EntityName { get; set; }
        public string FileNameFilter { get; set; }
        public string Tag { get; set; }
        public string Value { get; set; }

    }

    class ExcelReader
    {
        public static string excelFilePath = @"C:\Users\fasil.m\Desktop\SKUTA\Configs.xlsx";
        public static List<ExcelObject> PutExcelToList(string excelPath, string sheetName)
        {
            Application excelApp = new Application();

            Workbook excelBook = excelApp.Workbooks.Open(excelPath);
            _Worksheet excelSheet = excelBook.Sheets[sheetName];
            Range excelRange = excelSheet.UsedRange;
            List<ExcelObject> workSheetData = new List<ExcelObject>();

            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {
                ExcelObject rdc = new ExcelObject();

                if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                    rdc.EntityName = excelRange.Cells[i, 1].Value2.ToString();
                if (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                    rdc.FileNameFilter = excelRange.Cells[i, 2].Value2.ToString();
                if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                    rdc.Tag = excelRange.Cells[i, 3].Value2.ToString();
                if (excelRange.Cells[i, 4] != null && excelRange.Cells[i, 4].Value2 != null)
                    rdc.Value = excelRange.Cells[i, 4].Value2.ToString();
                workSheetData.Add(rdc);
            }

            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            return workSheetData;
        }


    }
}
