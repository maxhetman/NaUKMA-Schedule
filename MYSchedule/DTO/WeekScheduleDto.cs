using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MYSchedule.DTO
{
    public class WeekScheduleDto
    {
        public WeekDto WeekDto; //PPK FK
        public ScheduleRecordDto SheduleRecord; //PPK FK

        public override string ToString()
        {
            return $"{nameof(WeekDto)}: {WeekDto}, {nameof(SheduleRecord)}: {SheduleRecord}";
        }
    }
}
