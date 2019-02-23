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
        public static void Export(string header, DataTable dataTable)
        {
            var columnNames = Utils.Utils.GetColumnNames(dataTable);
            Application excel = new Application();

            excel.Application.Workbooks.Add(true);
            Worksheet worksheet = (Worksheet) excel.ActiveSheet;

            Utils.Utils.InitCommonStyle(worksheet);

            CreateHeader(header, worksheet, columnNames);
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
            var rowsCount = dataTable.Rows.Count+2;

            worksheet.Cells[1, 1].Font.Size = 15;
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 1].Borders.Weight = 2d;


            var headersRange = worksheet.Range[worksheet.Cells[2, 1],
                    worksheet.Cells[2, columnsCount]]
                .Cells;

            headersRange.Font.Size = 12;
            headersRange.Font.Bold = true;
            headersRange.Borders.Weight = 2d;

            //vertical lines
            for (int i = 1; i <= columnsCount; i++)
            {               
                worksheet.Range[worksheet.Cells[1, i], worksheet.Cells[rowsCount, i]].Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = 2d;
            }

            //horizontal lines
            for (int i = 1; i <= rowsCount; i++)
            {
                worksheet.Range[worksheet.Cells[i, 1], worksheet.Cells[i, columnsCount]].Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2d;
            }

            worksheet.Range["A1", "U500"].Columns.AutoFit();
            worksheet.Range["A1", "U500"].Rows.AutoFit();
        }

        private static void CreateHeader(string header, Worksheet worksheet, string[] columnNames)
        {
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, columnNames.Length]].Merge();
            worksheet.Cells[1, 1] = header;
            worksheet.Range[worksheet.Cells[1, 1],
                worksheet.Cells[1, 1]].Interior.Color = XlRgbColor.rgbBeige;

            int i=0;
            for ( i = 0; i < columnNames.Length; i++)
            {
                worksheet.Cells[2, i+1] = columnNames[i];
            }
            worksheet.Range[worksheet.Cells[2, 1],
                worksheet.Cells[2, i]].Interior.Color = XlRgbColor.rgbAntiqueWhite;
        }

        private static void FillData(Worksheet worksheet, DataTable dataTable)
        {            
            var headerOffset = 3;           
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
