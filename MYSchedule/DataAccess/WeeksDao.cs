using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
namespace MYSchedule.DataAccess
{
    public static class WeeksDao
    {
        //todo: distinct must not be here. fix this shit
        private const string getAllWeeks = "Select Number, Begin, [End] From [Week] Order By [Number]";

        public static DataTable GetAllWeeks()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = getAllWeeks;

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);
                return dataTable;
            }
        }
    }
}
