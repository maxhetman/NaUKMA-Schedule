using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
namespace MYSchedule.DataAccess
{
    public static class WeeksDao
    {
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

        public static string[] GetFormattedWeeks()
        {
            var weeksDt = GetAllWeeks();
            string[] res = new string[weeksDt.Rows.Count];
            for(int i = 0; i < res.Length; i++)
            {
                res[i] = Utils.Utils.GetFormattedWeek(weeksDt.Rows[i], false);
            }
            return res;
        }

        public static void SetFirstWeekDate(DateTime selectedDate)
        {
            var beginDate = selectedDate;
            var endDate = selectedDate.AddDays(7);
            for (int i = 1; i <= 15; i++)
            {
                //TSR
                if (i == 8)
                {
                    beginDate = beginDate.AddDays(7);
                    endDate = endDate.AddDays(7);
                    continue;
                }

                var beginDateStr = beginDate.ToString("MM/dd/yyyy");
                var endDateStr = endDate.ToString("MM/dd/yyyy");

                var query = "UPDATE [Week] SET [Begin] = \"" + beginDateStr + "\", [End] = \"" + endDateStr + "\"  WHERE Number = " + i;
                UpdateWeekInfo(query);
                beginDate = beginDate.AddDays(7);
                endDate = endDate.AddDays(7);
            }
        }


        private static void UpdateWeekInfo(string query)
        {
            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.UpdateCommand = new OleDbCommand();
                dataAdapter.UpdateCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.UpdateCommand.CommandType = CommandType.Text;
                dataAdapter.UpdateCommand.CommandText = query;

                dataAdapter.UpdateCommand.Connection.Open();
                dataAdapter.UpdateCommand.ExecuteNonQuery();
                dataAdapter.UpdateCommand.Connection.Close();

            }

        }
    }
}
