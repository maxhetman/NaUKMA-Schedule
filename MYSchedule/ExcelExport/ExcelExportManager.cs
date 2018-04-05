using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using MYSchedule.DTO;
using DataTable = System.Data.DataTable;
using Constants = MYSchedule.Utils.Constants;

namespace MYSchedule.ExcelExport
{
    public struct CellIndex
    {
        public int x;
        public int y;
    }

    public struct TeacherWeeksWithCellIndex
    {
        public string TeacherName;
        public List<int> Weeks;
        public CellIndex CellIndex;
        public string DayName;
    }

    class ExcelExportManager
    {

        #region Variables

        private static int finalXCoord;

        private static Dictionary<string, CellIndex> DayCellsIndexes = new Dictionary<string, CellIndex>()
        {
            {Constants.Monday, new CellIndex() {x = 2, y = 4}},
            {Constants.Tuesday, new CellIndex() {x = 2, y = 6}},
            {Constants.Wednesday, new CellIndex() {x = 2, y = 8}},
            {Constants.Thursday, new CellIndex() {x = 2, y = 10}},
            {Constants.Friday, new CellIndex() {x = 2, y = 12}},
            {Constants.Saturday, new CellIndex() {x = 2, y = 14}}
        };

        private static List<string> ClassRooms;

        #endregion

        public static void ShowAllClassRooms(DataTable dataTable)
        {
            Application excel = new Application();

            excel.Application.Workbooks.Add(true);
            Worksheet worksheet = (Worksheet) excel.ActiveSheet;

            InitStyle(worksheet);

            CreateHeader(worksheet);
            CreateSkeleton(worksheet, dataTable);
            FillClassRooms(worksheet);
            FillTable(dataTable, worksheet);

            excel.Visible = true;

            FinalStyleAdditions(worksheet);
           

            //   worksheet.Range["B1","B100"].EntireColumn.Style.Orientation = Microsoft.Office.Interop.Excel.XlOrientation.xlUpward;
            worksheet.Activate();
        }

       


        public static void SpecialityYearLesson(DataTable dataTable)
        {
            //var table = QueryManager.GetClassRoomsAvailability();

            //Application excel = new Application();

            //excel.Application.Workbooks.Add(true);

            //InitStyle(excel);


            //CreateShowClassRoomsExcelHeader(excel);
            //CreateShowAllClassRoomsExcelSkeleton(excel, table);


            //excel.Visible = true;
            //Worksheet worksheet = (Worksheet)excel.ActiveSheet;
            //worksheet.Activate();
        }

        #region Helpers

        private static void FillTable(DataTable dataTable, Worksheet worksheet)
        {
            List<TeacherWeeksWithCellIndex> teachersData = new List<TeacherWeeksWithCellIndex>();

            var currentTime = -1;

            foreach (DataRow row in dataTable.Rows)
            {

                var dayName = row[0].ToString();
                var lessonTimeNumber = LessonTimeDto.GetNumberFromPeriod(row[1].ToString());
                var classroom = row[2].ToString();
                var cellIndex = GetIndexByDayLessonClassroom(dayName, lessonTimeNumber, classroom);

                if (currentTime != lessonTimeNumber)
                {
                    teachersData.Clear(); //we dont give a fuck if teacher has same classes on different lesson times
                    currentTime = lessonTimeNumber;
                }

                var teacherName = string.Format("{0} {1}", row[3], row[4]);
                var weeks = row[5].ToString();

                AddCellInfo(worksheet, cellIndex, teacherName, weeks);

                TeacherWeeksWithCellIndex currentTeacherData = new TeacherWeeksWithCellIndex
                {
                    CellIndex = cellIndex,
                    TeacherName = teacherName,
                    Weeks = Utils.Utils.ParseWeeks(weeks),
                    DayName = dayName
                };

                foreach (var teacherData in teachersData)
                {
                    //teacher has more than one lesson in same time
                    if (teacherData.TeacherName == currentTeacherData.TeacherName
                        && teacherData.DayName == currentTeacherData.DayName
                        && teacherData.Weeks.Intersect(currentTeacherData.Weeks).Any())
                    {
                        SetCellBackground(worksheet, teacherData.CellIndex, 
                            new CellIndex { x= teacherData.CellIndex.x, y= teacherData.CellIndex.y + 1}, XlRgbColor.rgbIndianRed);
                        SetCellBackground(worksheet, currentTeacherData.CellIndex,
                            new CellIndex { x = currentTeacherData.CellIndex.x, y = currentTeacherData.CellIndex.y + 1 }, XlRgbColor.rgbIndianRed);
                    }
                }

                teachersData.Add(currentTeacherData);
            }
        }

        private static void AddCellInfo(Worksheet worksheet, CellIndex cellIndex, string teacherName, string weeks)
        {
            var prevWeeks = worksheet.Cells[cellIndex.x, cellIndex.y + 1].Text.ToString();

            var hasPrevWeeks = !string.IsNullOrEmpty(prevWeeks);

            if (!hasPrevWeeks)
            {
                worksheet.Cells[cellIndex.x, cellIndex.y] = teacherName;
                worksheet.Cells[cellIndex.x, cellIndex.y + 1] = weeks;
                return;
            }

            var prevName = worksheet.Cells[cellIndex.x, cellIndex.y].Text.ToString();
            worksheet.Cells[cellIndex.x, cellIndex.y] = prevName + "/\n" + teacherName;
            worksheet.Cells[cellIndex.x, cellIndex.y + 1] = prevWeeks + "/\n" + weeks;

            List<int> prevWeeksList = Utils.Utils.ParseWeeks(prevWeeks);
            List<int> currWeeksList = Utils.Utils.ParseWeeks(weeks);

            if (prevWeeksList.Intersect(currWeeksList).Any())
            {
                SetCellBackground(worksheet, cellIndex, new CellIndex {x = cellIndex.x, y = cellIndex.y + 1}, XlRgbColor.rgbRed);
            }
        }

        
        private static void SetCellBackground(Worksheet worksheet, CellIndex from, CellIndex to, XlRgbColor color)
        {
            worksheet.Range[worksheet.Cells[from.x, from.y],
                worksheet.Cells[to.x, to.y]].Interior.Color = color;
        }

        private static void CreateSkeleton(Worksheet worksheet, DataTable dataTable)
        {
            var classRoomNumbers = new HashSet<string>();

            foreach (DataRow row in dataTable.Rows)
            {
                var number = row[2].ToString();
                if (!classRoomNumbers.Contains(number))
                {
                    classRoomNumbers.Add(number);
                }
            }

            var len = classRoomNumbers.Count;

            for (int i = 0; i < 7; i++)
            {
                var start = 4 + i * len;
                worksheet.Range[worksheet.Cells[start, 1], worksheet.Cells[start + len - 1, 1]].Merge();
                worksheet.Cells[start, 1] = i + 1;
                worksheet.Range[worksheet.Cells[start, 2], worksheet.Cells[start + len - 1, 2]].Merge();
                worksheet.Cells[start, 2] = LessonTimeDto.GetPeriodFromNumber(i + 1);
            }

            ClassRooms = classRoomNumbers.ToList();
        }

        private static void FillClassRooms(Worksheet worksheet)
        {
            var classRoomLength = ClassRooms.Count;
            int j = 0;
            var start = 0;
            for (int i = 0; i < 7; i++)
            {
                 start = 4 + i * classRoomLength;

                for (j = 0; j < classRoomLength; j++)
                {
                    worksheet.Cells[start + j, 3] = ClassRooms[j];
                }
            }

            finalXCoord = start + j-1;
        }

        private static void InitStyle(Worksheet worksheet)
        {
            string startRange = "A1";
            string endRange = "U500";
            var currentRange = worksheet.Range[startRange, endRange];
            currentRange.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            currentRange.Style.VerticalAlignment = XlHAlign.xlHAlignCenter;
            currentRange.Style.NumberFormat = "@";

            // worksheet.Range["C1","C50"].Style.Orientation  = Microsoft.Office.Interop.Excel.XlOrientation.xlUpward;
        }

        private static void FinalStyleAdditions(Worksheet worksheet)
        {
            var classRoomLength = ClassRooms.Count;



            worksheet.get_Range("A1", "A1").Cells.Font.Size = 15;
            worksheet.get_Range("A1", "O3").Cells.Font.Bold = true;
            worksheet.get_Range("A1", "O3").Cells.Borders.Weight = 2d;
            worksheet.get_Range("A1", "C"+ finalXCoord).Cells.Borders.Weight = 2d;

            for (int i = 1; i <= 7; i++)
            {
                var xCoord = 3 + i*classRoomLength;
                worksheet.get_Range("A13", "O" + xCoord).Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2d;
            }

            for (int i = 0; i < 6; i++)
            {
                var yCoord = 5 + i * 2;
                worksheet.Range[worksheet.Cells[3, yCoord], worksheet.Cells[finalXCoord, yCoord]].Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = 2d;
            }
          

            worksheet.get_Range("O1", "O" + finalXCoord).Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = 2d;

            worksheet.Range["A1", "U500"].Columns.AutoFit();
            worksheet.Range["A1", "U500"].Rows.AutoFit();
        }

        private static CellIndex GetIndexByDayLessonClassroom(string dayName, int lessonNumber, string classroom)
        {
            var len = ClassRooms.Count;
            var x = 4 + (lessonNumber - 1) * len + (ClassRooms.IndexOf(classroom));
            var y = DayCellsIndexes[dayName].y;
            return new CellIndex {x = x, y = y};
        }

        private static void CreateHeader(Worksheet worksheet)
        {
            var header = "Розклад аудиторій на весну 2017-2018";

            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 15]].Merge();
            worksheet.Cells[1, 1] = header;


            worksheet.Range[worksheet.Cells[2, 4], worksheet.Cells[2, 5]].Merge();
            worksheet.Range[worksheet.Cells[2, 6], worksheet.Cells[2, 7]].Merge();
            worksheet.Range[worksheet.Cells[2, 8], worksheet.Cells[2, 9]].Merge();
            worksheet.Range[worksheet.Cells[2, 10], worksheet.Cells[2, 11]].Merge();
            worksheet.Range[worksheet.Cells[2, 12], worksheet.Cells[2, 13]].Merge();
            worksheet.Range[worksheet.Cells[2, 14], worksheet.Cells[2, 15]].Merge();

            worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[3, 1]].Merge();
            worksheet.Range[worksheet.Cells[2, 2], worksheet.Cells[3, 2]].Merge();
            worksheet.Range[worksheet.Cells[2, 3], worksheet.Cells[3, 3]].Merge();

            worksheet.Cells[2, 1] = "№";
            worksheet.Cells[2, 2] = "Час";
            worksheet.Cells[2, 3] = "Ауд";
            worksheet.Cells[2, 4] = Constants.Monday;
            worksheet.Cells[2, 6] = Constants.Tuesday;
            worksheet.Cells[2, 8] = Constants.Wednesday;
            worksheet.Cells[2, 10] = Constants.Thursday;
            worksheet.Cells[2, 12] = Constants.Friday;
            worksheet.Cells[2, 14] = Constants.Saturday;

            for (int i = 4; i <= 14; i += 2)
            {
                worksheet.Cells[3, i] = "Прізвище";
                worksheet.Cells[3, i + 1] = "№ т.";
            }
        }

        #endregion


    }
}
