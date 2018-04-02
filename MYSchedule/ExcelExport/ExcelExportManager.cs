using System.Collections.Generic;
using System.Data;
using Microsoft.Office.Interop.Excel;
using MYSchedule.DataAccess;
using DataTable = System.Data.DataTable;

namespace MYSchedule.ExcelExport
{
    class ExcelExportManager
    {
        public static void ShowAllClassRooms(DataTable dataTable)
        {
            var table = QueryManager.GetClassRoomsBusyness();

           Application excel = new Application();

            excel.Application.Workbooks.Add(true);

            int columnIndex = 0;

            AlignTextHorizontal(excel);


            CreateExcelHeader(excel);

            CreateExcelSkeleton(excel, dataTable);

            int rowIndex = 0;

            //foreach (DataRow row in table.Rows)
            //{
            //    rowIndex++;
            //    columnIndex = 0;
            //    foreach (DataColumn col in table.Columns)
            //    {
            //        columnIndex++;
            //        if (columnIndex == 4 || columnIndex == 5 || columnIndex == 6)
            //        {
            //            if (columnIndex == 4)
            //            {
            //                excel.Cells[rowIndex + 1, columnIndex]
            //                    = Enum.GetName(typeof(Occupation), row[col.ColumnName]);
            //            }

            //            if (columnIndex == 5)
            //            {
            //                excel.Cells[rowIndex + 1, columnIndex]
            //                    = Enum.GetName(typeof(MaritalStatus), row[col.ColumnName]);
            //            }

            //            if (columnIndex == 6)
            //            {
            //                excel.Cells[rowIndex + 1, columnIndex]
            //                    = Enum.GetName(typeof(HealthStatus), row[col.ColumnName]);
            //            }
            //        }
            //        else
            //        {
            //            excel.Cells[rowIndex + 1, columnIndex] = row[col.ColumnName].ToString();
            //        }
            //    }
            //}

            excel.Visible = true;
            Worksheet worksheet = (Worksheet) excel.ActiveSheet;
            worksheet.Activate();
        }

        private static void CreateExcelSkeleton(Application excel, DataTable dataTable)
        {
            var classRoomNumbers = new HashSet<string>();

            foreach (DataRow row in dataTable.Rows)
            {
                var number = row[3].ToString();
                if (!classRoomNumbers.Contains(number))
                {
                    classRoomNumbers.Add(number);
                }
            }

            var len = classRoomNumbers.Count;
            for (int i = 0; i < 7; i++)
            {
                var start = 4 + i * len;
                excel.Range[excel.Cells[start, 1], excel.Cells[start + len - 1, 1]].Merge();
                excel.Range[excel.Cells[start, 2], excel.Cells[start + len - 1, 2]].Merge();
            }
        }

        private static void AlignTextHorizontal(Application excel)
        {
            string startRange = "A1";
            string endRange = "U500";
            var currentRange = excel.Range[startRange, endRange];
            currentRange.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        private static void CreateExcelHeader(Application excel)
        {
            var header = "Розклад аудиторій на весну 2017-2018";

            excel.Range[excel.Cells[1, 1], excel.Cells[1, 15]].Merge();
            excel.Cells[1, 1] = header;




            excel.Range[excel.Cells[2, 4], excel.Cells[2, 5]].Merge();
            excel.Range[excel.Cells[2, 6], excel.Cells[2, 7]].Merge();
            excel.Range[excel.Cells[2, 8], excel.Cells[2, 9]].Merge();
            excel.Range[excel.Cells[2, 10], excel.Cells[2, 11]].Merge();
            excel.Range[excel.Cells[2, 12], excel.Cells[2, 13]].Merge();
            excel.Range[excel.Cells[2, 14], excel.Cells[2, 15]].Merge();

            excel.Range[excel.Cells[2, 1], excel.Cells[3, 1]].Merge();
            excel.Range[excel.Cells[2, 2], excel.Cells[3, 2]].Merge();
            excel.Range[excel.Cells[2, 3], excel.Cells[3, 3]].Merge();

            excel.Cells[2, 1] = "№";
            excel.Cells[2, 2] = "Час";
            excel.Cells[2, 3] = "Ауд";
            excel.Cells[2, 4] = "Понеділок";
            excel.Cells[2, 6] = "Вівторок";
            excel.Cells[2, 8] = "Середа";
            excel.Cells[2, 10] = "Четвер";
            excel.Cells[2, 12] = "П`ятниця";
            excel.Cells[2, 14] = "Субота";

            for (int i = 4; i <= 14; i+=2 )
            {
                excel.Cells[3, i] = "Прізвище";
                excel.Cells[3, i+1] = "№ т.";
            }
        }
    }
}
