namespace ExcelDataExtractor;

public class WorkedDay
{
    public int DayNumber { get; set; }
    public int Norm { get; set; }
    public double NightHours { get; set; }
    public double DayHours { get; set; }
    public DayTypeEnum DayType { get; set; }
    public DateTime Date { get; set; }
}