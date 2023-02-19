namespace ExcelDataExtractor;

public class ExcelService
{
    private List<EmployeeModel> _employees { get; set; }
    
    public ExcelService(List<EmployeeModel> employees)
    {
        _employees = employees;
    }

    public void UpdateEmployeesData()
    {
        foreach (var employee in _employees)
        {
            double overTime = 0.0;
            int consecutiveWorkedDays = 0;
            
            int index = 0;
            foreach (var day in employee.WorkedDays)
            {
                if (day.DayHours + day.NightHours > day.Norm)
                {
                    overTime += (day.DayHours + day.NightHours) - day.Norm;
                }

                DealWithConsecutiveDays(consecutiveWorkedDays, day, employee, index, overTime);
                
                index++;
            }

            var freeDays = employee.WorkedDays.Where(x => x.DayType == DayTypeEnum.Concediu);
            DealWithOverTime(overTime, freeDays, employee);
        }
    }

    private void DealWithConsecutiveDays(int consecutiveWorkedDays, WorkedDay day, EmployeeModel employee, int index, double overTime)
    {
        if (consecutiveWorkedDays == 0)
        {
            consecutiveWorkedDays++;
        }
        else
        {
            if (day.DayNumber != employee.WorkedDays[index - 1].DayNumber + 1)
            {
                consecutiveWorkedDays = 0;
            }

            consecutiveWorkedDays++;
        }

        if (consecutiveWorkedDays >= 7)
        {
            var nextDay = employee.WorkedDays[index + 1];
            overTime += day.DayHours + day.NightHours + 
                        nextDay.DayHours + nextDay.NightHours;
                    
            day.DayHours = 0;
            day.NightHours = 0;
            nextDay.DayHours = 0;
            nextDay.NightHours = 0;
            day.DayType = DayTypeEnum.Concediu;
            nextDay.DayType = DayTypeEnum.Concediu;
        }
    }

    private void DealWithOverTime(double overTime, IEnumerable<WorkedDay> freeDays, EmployeeModel employee)
    {
        int index = 0;
        while (overTime > 0 && index < freeDays.Count())
        {
            var day = employee.WorkedDays[index];
            if (overTime > 16)
            {
                day.DayHours = 8;
                day.NightHours = 8;
                overTime -= 16;
            }
            else if (overTime < 16 && overTime > 8)
            {
                day.DayHours = 8;
                day.NightHours = overTime - 8;
            }
            else if (overTime < 8)
            {
                day.DayHours = overTime;
            }
            index++;
        }
    }
}