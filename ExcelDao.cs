using System.Runtime.InteropServices;

namespace ExcelDataExtractor;
using Microsoft.Office.Interop.Excel;

public class ExcelDao
{
    private Application _excelApp = new Application();
    private Workbook _workbook;
    private Worksheet _worksheet;
    private Range _xlRange;

    public ExcelDao(string path)
    {
        _workbook = _excelApp.Workbooks.Open(path);
        _worksheet = _workbook.Sheets[1];
        _xlRange = _worksheet.UsedRange;
    }

    public List<EmployeeModel> GetAllEmployees()
    {
        List<EmployeeModel> employees = new List<EmployeeModel>();
        
        for (int i = 2; i < _xlRange.Rows.Count; i++)
        {
            string mark = _xlRange.Cells[i, 1].Value2.ToString();
            string name = _xlRange.Cells[i, 2].Value2.ToString();
            string id = _xlRange.Cells[i, 3].Value2.ToString();
            
            int norm = Int32.Parse(_xlRange.Cells[i, 4].Value2.ToString());
            double nightHours = Double.Parse(_xlRange.Cells[i, 5].Value2.ToString());
            double dayHours = Double.Parse(_xlRange.Cells[i, 6].Value2.ToString());
            Enum.TryParse(_xlRange.Cells[i, 7].Value2.ToString(), out DayTypeEnum dayType);
            string dateAsString = _xlRange.Cells[i, 8].Value.ToString();
            DateTime date = DateTime.Parse(dateAsString);
            int dayNumber = date.Day;
            
            var workedDay = new WorkedDay
            {
                DayNumber = dayNumber,
                Norm = norm,
                NightHours = nightHours,
                DayHours = dayHours,
                DayType = dayType,
                Date = date
            };
            
            if (!employees.Select(x => x.Name).Contains(name))
            {
                var employee = new EmployeeModel();
                employee.Name = name;
                employee.Mark = mark;
                employee.Id = id;
                employee.WorkedDays.Add(workedDay);
                
                //add employee to return list
                employees.Add(employee);
            }
            else
            {
                var employeeToBeUpdated = employees.Single(x => x.Name == name);
                employeeToBeUpdated.WorkedDays.Add(workedDay);
            }
            
        }

        return employees;
    }

    public void ExportData(List<EmployeeModel> employees)
    {
        var tableSize = "|{0,10}|{1,15}|{2,35}|{3,10}|{4,10}|{5,10}|{6,10}|{7,25}|";
        var topRow = String.Format(tableSize,
            "Marca", "Nume", "IdAngajat", "Norma", "Ore noapte", "Ore zi", "Tip ore", "Ziua");
        Console.WriteLine(topRow);
        for (int i = 0; i < employees.Count; i++)
        {
            var workedDays = employees[i].WorkedDays;
            for (int y = 0; y < workedDays.Count; y++)
            {
                var row = String.Format(tableSize,
                    employees[i].Mark, employees[i].Name, employees[i].Id, workedDays[y].Norm, workedDays[y].NightHours,
                    workedDays[y].DayHours, workedDays[y].DayType.ToString(), workedDays[y].Date);
                var rowSeparator = String.Format(tableSize,
                    new String('-', 10), new String('-', 15), new String('-', 35), new String('-', 10), new String('-', 10), new String('-', 10), new String('-', 10), new String('-', 25));
                Console.WriteLine(rowSeparator);
                Console.WriteLine(row);
            }
        }
    }

    public void CloseExcelApp()
    {
        Marshal.ReleaseComObject(_xlRange);
        Marshal.ReleaseComObject(_worksheet);

        //close and release
        _workbook.Close();
        Marshal.ReleaseComObject(_workbook);

        //quit and release
        _excelApp.Quit();
        Marshal.ReleaseComObject(_excelApp);
    }
}