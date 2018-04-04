using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MYSchedule.Utils
{
    public static class Utils
    {

        public static List<int> ParseWeeks(string weeks)
        {
            weeks = weeks.Replace(" ", String.Empty);
            var weeksArray = weeks.Split(',');
            var weeksList = new List<int>();
            foreach (var value in weeksArray)
            {
                int weekNumber;
                int.TryParse(value, out weekNumber);

                if (weekNumber > 0)
                {
                    weeksList.Add(weekNumber);
                    continue;
                }
                var notNumberRegexp = @"[^\d]";
                var leftMatch = Regex.Match(value, notNumberRegexp);
                var rightMatch = Regex.Match(value, notNumberRegexp, RegexOptions.RightToLeft);

                if (leftMatch.Success && rightMatch.Success)
                {
                    var leftNumber = int.Parse(value.Substring(0, leftMatch.Index));
                    var rightNumber = int.Parse(value.Substring(rightMatch.Index + 1));

                    for (int i = leftNumber; i <= rightNumber; i++)
                        weeksList.Add(i);
                }
            }
            return weeksList;
        }
    }
}
