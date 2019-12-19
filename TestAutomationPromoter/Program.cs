using System;
using System.Configuration;

namespace TestAutomationPromoter
{
    class Program
    {
        static void Main(string[] args)

        {
            string excelFilePath = ConfigurationManager.AppSettings["ExcelPath"];
            string excelsheetName = ConfigurationManager.AppSettings["ExcelSheetName"];


            var collection = ExcelReader.PutExcelToList(excelFilePath, excelsheetName);

            var spath = @"C:\\Users\\fasil.m\\Desktop\\SKUTA\\TestAutomation\\";

            DataSetHelper.PromoteDataSet(spath, collection);

            Console.ReadKey();

        }
    }
}
