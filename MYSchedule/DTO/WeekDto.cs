using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MYSchedule.DTO
{
    public class WeekDto
    {
        public int Number;
        public DateTime Begin;
        public DateTime End;

        public override string ToString()
        {
            return $"{nameof(Number)}: {Number}, {nameof(Begin)}: {Begin}, {nameof(End)}: {End}";
        }
    }
}
