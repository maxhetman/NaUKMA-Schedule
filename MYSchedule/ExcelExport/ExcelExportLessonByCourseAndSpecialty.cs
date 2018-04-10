using System;
using System.Collections.Generic;
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
        #region Variables

        private static Dictionary<string, CellIndex> WeekNumberCellIndex = new Dictionary<string, CellIndex>();

        #endregion

        public static void LessonScheduleByCourseAndSpecialty
            (DataTable dataTable)
        {
            Application excel = new Application();

            excel.Application.Workbooks.Add(true);
            Worksheet worksheet = (Worksheet)excel.ActiveSheet;

            Utils.Utils.InitCommonStyle(worksheet);

            CreateHeader(worksheet);
            CreateSkeleton(worksheet);
            FillData(worksheet, dataTable);

            FinalStyleAdditions(worksheet);

            excel.Visible = true;
            worksheet.Activate();
        }

        #region Helpers

        private static void CreateSkeleton(Worksheet worksheet)
        {
            var weeks = WeeksDao.GetAllWeeks();
            var startIndex = 4;

            foreach (DataRow week in weeks.Rows)
            {
                worksheet.Cells[3, startIndex] = FormattedWeekPeriod(week);
                WeekNumberCellIndex.Add(week[0].ToString(), new CellIndex(3,startIndex));
                startIndex++;
            }
        }

        private static string FormattedWeekPeriod(DataRow week)
        {
            var start = Convert.ToDateTime(week[1]).Date;
            var endDate = Convert.ToDateTime(week[2]).Date;
            return string.Format("{0} т. \n {1}.{2} - {3}.{4}", week[0], start.Day, start.Month, endDate.Day, endDate.Month);
        }

        private static void CreateHeader(Worksheet worksheet)
        {
            var header = "МП \"Комп`ютерні науки\", 1 курс, Англійська мова";

            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 17]].Merge();
            worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[2, 17]].Merge();
            worksheet.Cells[1, 1] = header;
            worksheet.Cells[2, 1] = "Тижні";




            worksheet.Cells[3, 1] = "День\nЧас";
            worksheet.Cells[3, 2] = "Ауд.";
            worksheet.Cells[3, 3] = "Викладач";
        }

        private static void FillData(Worksheet worksheet, DataTable dataTable)
        {
            foreach (DataRow one in dataTable.Rows)
            {
                Console.WriteLine($"{one[0]}, {one[1]}, {one[2]}, {one[3]}, {one[4]}, {one[5]}, {one[6]}");
            }
        }

        private static void FinalStyleAdditions(Worksheet worksheet)
        {
            worksheet.Range["A1", "A1"].Cells.Font.Size = 15;
            worksheet.Range["A1", "Q3"].Cells.Font.Bold = true;
            worksheet.Range["A1", "Q3"].Cells.Borders.Weight = 2d;

            //for (int i = 1; i <= 7; i++)
            //{
            //    var xCoord = 3 + i * classRoomLength;
            //    worksheet.Range["A13", "O" + xCoord].Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2d;
            //}

            //for (int i = 0; i < 6; i++)
            //{
            //    var yCoord = 5 + i * 2;
            //    worksheet.Range[worksheet.Cells[3, yCoord], worksheet.Cells[finalXCoord, yCoord]].Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = 2d;
            //}

            worksheet.Range["A1", "U500"].Columns.AutoFit();
            worksheet.Range["A1", "U500"].Rows.AutoFit();
        }
        #endregion
    }
}
