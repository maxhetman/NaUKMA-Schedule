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

        private static int lastWeekYIndex;
        private static int lastXIndex;

        #endregion

        public static void LessonScheduleByCourseAndSpecialty
            (string header, DataTable dataTable)
        {
            WeekNumberCellIndex.Clear();
            Application excel = new Application();

            excel.Application.Workbooks.Add(true);
            Worksheet worksheet = (Worksheet)excel.ActiveSheet;

            Utils.Utils.InitCommonStyle(worksheet);

            CreateHeader(header, worksheet);
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
            lastWeekYIndex = startIndex;
        }

        private static string FormattedWeekPeriod(DataRow week)
        {
            var start = Convert.ToDateTime(week[1]).Date;
            var endDate = Convert.ToDateTime(week[2]).Date;
            return string.Format("{0} т. \n {1}.{2} - {3}.{4}", week[0], start.Day.ToString("00"), start.Month.ToString("00"), endDate.Day.ToString("00"), endDate.Month.ToString("00"));
        }

        private static void CreateHeader(string header, Worksheet worksheet)
        {
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 17]].Merge();
            worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[2, 17]].Merge();
            worksheet.Cells[1, 1] = header;
            worksheet.Range[worksheet.Cells[1, 1],
                worksheet.Cells[1, 1]].Interior.Color = XlRgbColor.rgbBeige;

            worksheet.Cells[2, 1] = "Тижні";

            worksheet.Cells[3, 1] = "День\nЧас";
            worksheet.Cells[3, 2] = "Ауд.";
            worksheet.Cells[3, 3] = "Викладач";
        }

        private static void FillData(Worksheet worksheet, DataTable dataTable)
        {
            var currentDayName = string.Empty;
            int currentLessonTimeNumber=-1;
            string currentLessonTime = string.Empty;
            var currentDayTimeCell = new CellIndex(4, 1);


            string currentClassRoom = string.Empty;
            var currentClassRoomCell = new CellIndex(4, 2);


            string currerntTeacher = string.Empty;
            var currentTeacherCell = new CellIndex(4, 3);

            foreach (DataRow dataRow in dataTable.Rows)
            {
                var dayTimeWasChanged = false;
                var dayName = dataRow[1].ToString();
                var lessonTimeNumber = int.Parse(dataRow[2].ToString());

                if (string.IsNullOrEmpty(currentDayName) || currentDayName != dayName)
                {
                    dayTimeWasChanged = true;
                    currentDayName = dayName;
                }

                if (currentLessonTimeNumber != lessonTimeNumber)
                {
                    dayTimeWasChanged = true;
                    currentLessonTimeNumber = lessonTimeNumber;
                    currentLessonTime = LessonTimeDto.GetPeriodFromNumber(lessonTimeNumber);
                }

                if (dayTimeWasChanged)
                {
                    worksheet.Range[worksheet.Cells[currentDayTimeCell.x, currentDayTimeCell.y],
                        worksheet.Cells[currentDayTimeCell.x, lastWeekYIndex-1]].Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                    worksheet.Cells[currentDayTimeCell.x, currentDayTimeCell.y] = currentDayName + "\n" + currentLessonTime;
                    currentDayTimeCell.x++;
                }


                var classRoom = dataRow[3].ToString();

                if (currentClassRoom != classRoom || dayTimeWasChanged)
                {
                    worksheet.Range[worksheet.Cells[currentClassRoomCell.x, currentClassRoomCell.y],
                        worksheet.Cells[currentClassRoomCell.x, lastWeekYIndex-1]].Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                    if (currentClassRoomCell.x % 2 == 0)
                    {
                        worksheet.Range[worksheet.Cells[currentClassRoomCell.x, currentClassRoomCell.y],
                                worksheet.Cells[currentClassRoomCell.x, lastWeekYIndex - 1]]
                            .Interior.Color = XlRgbColor.rgbLightGray;
                    }

                   

                    currentClassRoom = classRoom;
                    worksheet.Cells[currentClassRoomCell.x, currentClassRoomCell.y] = currentClassRoom;
                    currentClassRoomCell.x++;

                    // if dayTime Cell was not changed we need to merge cells
                    if (!dayTimeWasChanged)
                    {
                         var previousXCoord = currentDayTimeCell.x - 1;
                        worksheet.Range[worksheet.Cells[previousXCoord, currentDayTimeCell.y],
                            worksheet.Cells[currentDayTimeCell.x++, currentDayTimeCell.y]].Merge();
                    }
                }

                var teacher = dataRow[4].ToString()+ " " + dataRow[5].ToString();


                if (currerntTeacher != teacher || dayTimeWasChanged)
                {
                    currerntTeacher = teacher;
                    worksheet.Cells[currentTeacherCell.x, currentTeacherCell.y] = currerntTeacher;
                    lastXIndex = currentTeacherCell.x;
                    currentTeacherCell.x++;
                }

                var weekNumber = dataRow[0].ToString();
                var weekXCoord = currentTeacherCell.x-1;
                var weekYCoord = WeekNumberCellIndex[weekNumber].y;

                worksheet.Cells[weekXCoord, weekYCoord] = dataRow[6].ToString() + " "+ dataRow[7].ToString();

               // Console.WriteLine($"{dataRow[0]}, {dataRow[1]}, {dataRow[2]}, {dataRow[3]}, {dataRow[4]}, {dataRow[5]}, {dataRow[6]}");

                dayTimeWasChanged = false;
            }
        }

        private static void FinalStyleAdditions(Worksheet worksheet)
        {
            worksheet.Range["A1", "A1"].Cells.Font.Size = 15;
            worksheet.Range["A2", "A2"].Cells.Font.Size = 12;
            worksheet.Range["A1", "Q3"].Cells.Font.Bold = true;
            worksheet.Range["A1", "Q3"].Cells.Borders.Weight = 2d;

            for (int i = 1; i < lastWeekYIndex; i++)
            {
                worksheet.Range[worksheet.Cells[3, i], worksheet.Cells[lastXIndex, i]].Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = 2d;
            }

            worksheet.Range[worksheet.Cells[3, 4],
                worksheet.Cells[3, lastWeekYIndex - 1]].Interior.Color = XlRgbColor.rgbAntiqueWhite;

            worksheet.Range["A1", "U500"].Columns.AutoFit();
            worksheet.Range["A1", "U500"].Rows.AutoFit();
            worksheet.Columns[1].ColumnWidth = 14;
            worksheet.Columns[2].ColumnWidth = 10;

            //for (int i = 4; i < 8; i++)
            //{
            //    worksheet.Rows[i].RowHeight= 20;
            //}

            //for (int i = 4; i < lastWeekYIndex; i++)
            //{
            //    worksheet.Columns[i].ColumnWidth= 9;
            //}

        }
        #endregion
    }
}
