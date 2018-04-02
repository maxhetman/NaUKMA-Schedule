using System;
using MYSchedule.Utils;

namespace MYSchedule.DTO
{
    public class LessonTimeDto
    {
        public int Number; //1,2,..,7
        public string LessonTimePeriod; //8:30-9:50...

        public static int GetNumberFromPeriod(string period)
        {
            period = period.Replace(" ", String.Empty);

            switch (period)
            {
                case "8:30-9:50":
                    return 1;
                case "10:00-11:20":
                    return 2;
                case "11:40-13:00":
                    return 3;
                case "13:30-14:50":
                    return 4;
                case "15:00-16:20":
                    return 5;
                case "16:30-17:50":
                    return 6;
                case "18:00-19:20":
                    return 7;
                default:
                    Logger.LogException("[LessonDto] Wrong LessonTimePeriod: " + period);
                    return -1;
            }        
            
;        }

        public override string ToString()
        {
            return $"{nameof(Number)}: {Number}, {nameof(LessonTimePeriod)}: {LessonTimePeriod}";
        }
    }
}
