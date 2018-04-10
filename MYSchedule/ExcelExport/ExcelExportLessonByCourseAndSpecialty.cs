using System;
using System.Data;
using Microsoft.Office.Interop.Excel;
using MYSchedule.DataAccess;
using MYSchedule.DTO;
using Constants = MYSchedule.Utils.Constants;
using DataTable = System.Data.DataTable;

namespace MYSchedule.ExcelExport
{
    public class ExcelExportLessonByCourseAndSpecialty
    {
        public static void LessonScheduleByCourseAndSpecialty
            (DataTable dataTable)
        {
            Application excel = new Application();

            excel.Application.Workbooks.Add(true);
            Worksheet worksheet = (Worksheet)excel.ActiveSheet;

            Utils.Utils.InitCommonStyle(worksheet);

            CreateHeader(worksheet);
            CreateSkeleton(worksheet, dataTable);
            //FillClassRooms(worksheet);
            //FillTable(dataTable, worksheet);

            excel.Visible = true;

            //FinalStyleAdditions(worksheet);


            //   worksheet.Range["B1","B100"].EntireColumn.Style.Orientation = Microsoft.Office.Interop.Excel.XlOrientation.xlUpward;
            worksheet.Activate();
        }

        private static void CreateSkeleton(Worksheet worksheet, DataTable dataTable)
        {
            var weeks = WeeksDao.GetAllWeeks();

            foreach (DataRow week in dataTable.Rows)
            {
                Console.WriteLine($"{week[0]}, {week[1]}, {week[2]}");

            }

            var startIndex = 3;

            //foreach (DataRow week in dataTable.Rows)
            //{
            //    worksheet.Cells[3, startIndex] = FormattedWeekPeriod(week);
            //    startIndex++;
            //}
            for (int i = 1; i <= 15; i++)
            {
                if (i == 8)
                {
                    startIndex -= 1;
                    continue;
                }

                worksheet.Cells[3, i + startIndex] = i + " т. \n (15.01-21.01)";
            }

            

        }

        private static string FormattedWeekPeriod(DataRow week)
        {
            var start = Convert.ToDateTime(week[1]).Date;
            var endDate = Convert.ToDateTime(week[2]).Date;
            return string.Format("{0} т. \n {1} - {2}", week[0], start, endDate);
        }

        private static void CreateHeader(Worksheet worksheet)
        {
            var header = "МП \"Комп`ютерні науки\", 1 курс, Англійська мова";

            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 17]].Merge();
            worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[2, 17]].Merge();
            worksheet.Cells[1, 1] = header;
            worksheet.Cells[2, 1] = "Тижні";


            //worksheet.Range[worksheet.Cells[2, 4], worksheet.Cells[2, 5]].Merge();
            //worksheet.Range[worksheet.Cells[2, 6], worksheet.Cells[2, 7]].Merge();
            //worksheet.Range[worksheet.Cells[2, 8], worksheet.Cells[2, 9]].Merge();
            //worksheet.Range[worksheet.Cells[2, 10], worksheet.Cells[2, 11]].Merge();
            //worksheet.Range[worksheet.Cells[2, 12], worksheet.Cells[2, 13]].Merge();
            //worksheet.Range[worksheet.Cells[2, 14], worksheet.Cells[2, 15]].Merge();

            //worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[3, 1]].Merge();
            //worksheet.Range[worksheet.Cells[2, 2], worksheet.Cells[3, 2]].Merge();
            //worksheet.Range[worksheet.Cells[2, 3], worksheet.Cells[3, 3]].Merge();

            worksheet.Cells[3,1] = "День\nЧас";
            worksheet.Cells[3, 2] = "Ауд.";
            worksheet.Cells[3, 3] = "Викладач";
        }

    }
}
