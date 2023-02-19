

namespace ExcelDataExtractor
{
    class Program
    {

        static void Main(string[] args)
        {
            const string EXCEL_PATH = "C:\\Alex\\Diverse\\Coding\\ConsoleApp1\\ExcelData\\StructuraDateTest.xlsx";
            
            //import data
            var excelDao = new ExcelDao(EXCEL_PATH);
            var employees = excelDao.GetAllEmployees();
            
            //update data
            var excelService = new ExcelService(employees);
            excelService.UpdateEmployeesData();
            
            excelDao.CloseExcelApp();
            
            //export data
            excelDao.ExportData(employees);
        }
    }
}