using System;
using System.Data;
using System.Data.OleDb;
using MYSchedule.Utils;

namespace MYSchedule.DataAccess
{
    public class WeekScheduleDao
    {
        private const string insertWeekSchedule = "Insert Into WeekSchedule(WeekNumber, ScheduleRecordId)" +
                                             " Values (@WeekNumber, @ScheduleRecordId)";

        public static void AddWeekSchedule(int weekNumber, int scheduleRecordId)
        {
            try
            {
                using (OleDbCommand oleDbCommand = new OleDbCommand())
                {
                    // Set the command object properties
                    oleDbCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                    oleDbCommand.CommandType = CommandType.Text;
                    oleDbCommand.CommandText = insertWeekSchedule;

                    // Add the input parameters to the parameter collection
                    oleDbCommand.Parameters.AddWithValue("@WeekNumber", weekNumber);
                    oleDbCommand.Parameters.AddWithValue("@ScheduleRecordId", scheduleRecordId);
                    // Open the connection, execute the query and close the connection
                    oleDbCommand.Connection.Open();
                    oleDbCommand.ExecuteNonQuery();
                    oleDbCommand.Connection.Close();
                }
            }   
            catch (Exception e)
            {
                Logger.LogException(e);
                Console.WriteLine("Exception in adding week schedule");
            }
        }
    }
}
