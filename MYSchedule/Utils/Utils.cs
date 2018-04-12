using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace MYSchedule.Utils
{
    public static class Utils
    {


        public static void InitCommonStyle(Worksheet worksheet)
        {
            string startRange = "A1";
            string endRange = "U500";
            var currentRange = worksheet.Range[startRange, endRange];
            currentRange.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            currentRange.Style.VerticalAlignment = XlHAlign.xlHAlignCenter;
            currentRange.Style.NumberFormat = "@";

            // worksheet.Range["C1","C50"].Style.Orientation  = Microsoft.Office.Interop.Excel.XlOrientation.xlUpward;
        }

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

        public static string[] GetColumnNames(DataTable dataTable)
        {
            string[] res = new string[dataTable.Columns.Count];
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                res[i] = dataTable.Columns[i].ColumnName.ToString();
            }
            return res;
        }
    }
}
