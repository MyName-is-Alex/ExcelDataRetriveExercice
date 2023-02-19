namespace ExcelDataExtractor;

public class EmployeeModel
{
    public string Mark { get; set; }
    public string Name { get; set; }
    public string Id { get; set; }
    public List<WorkedDay> WorkedDays { get; set; }

    public EmployeeModel()
    {
        WorkedDays = new List<WorkedDay>();
    }
}