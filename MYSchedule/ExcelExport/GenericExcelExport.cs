using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace MYSchedule.ExcelExport
{
    public static class GenericExcelExport
    {
        public static void Export(string[] columnNames, DataTable dataTable)
        {
            Application excel = new Application();

            excel.Application.Workbooks.Add(true);
            Worksheet worksheet = (Worksheet) excel.ActiveSheet;

            Utils.Utils.InitCommonStyle(worksheet);

            CreateHeader(worksheet, columnNames);
            FillData(worksheet, dataTable);
            FinalStyleAdditions(worksheet, columnNames, dataTable);
            excel.Visible = true;

            worksheet.Activate();
        }

        public static void SetCellBackground(Worksheet worksheet, CellIndex from, CellIndex to, XlRgbColor color)
        {
            worksheet.Range[worksheet.Cells[from.x, from.y],
                worksheet.Cells[to.x, to.y]].Interior.Color = color;
        }

        private static void FinalStyleAdditions(Worksheet worksheet, string[] columnNames, DataTable dataTable)
        {

            var columnsCount = dataTable.Columns.Count;
            var rowsCount = dataTable.Rows.Count;

            var headersRange = worksheet.Range[worksheet.Cells[1, 1],
                    worksheet.Cells[1, columnsCount]]
                .Cells;

            headersRange.Font.Size = 15;
            headersRange.Font.Bold = true;
            headersRange.Borders.Weight = 2d;

            //vertical lines
            for (int i = 1; i <= columnsCount; i++)
            {               
                worksheet.Range[worksheet.Cells[1, i], worksheet.Cells[rowsCount + 1, i]].Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = 2d;
            }

            //horizontal lines
            for (int i = 1; i <= rowsCount + 1; i++)
            {
                worksheet.Range[worksheet.Cells[i, 1], worksheet.Cells[i, columnsCount]].Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2d;
            }

            worksheet.Range["A1", "U500"].Columns.AutoFit();
            worksheet.Range["A1", "U500"].Rows.AutoFit();
        }

        private static void CreateHeader(Worksheet worksheet, string[] columnNames)
        {
            for (int i = 0; i < columnNames.Length; i++)
            {
                worksheet.Cells[1, i+1] = columnNames[i];
            }
        }

        private static void FillData(Worksheet worksheet, DataTable dataTable)
        {            
            var headerOffset = 2;           
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                var currentRow = dataTable.Rows[i];
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    worksheet.Cells[headerOffset + i, 1 + j] = currentRow[j];
                }
            }
        }
    }
}
