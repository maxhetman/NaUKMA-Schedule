using System.Data;
using System.Data.OleDb;

namespace MYSchedule.DataAccess
{
    public static class DBAccessManager
    {
        private const string clearWeekScheduleQuery = "Delete * From WeekSchedule";
        private const string clearScheduleRecordQuery = "Delete * From ScheduleRecord";
        private const string clearTeacherQuery = "Delete * From Teacher";
        private const string classRoomInconsistenseQuery = "SELECT DayName, LessonTimePeriod, ClassRoomNumber, WS.WeekNumber AS [Week] FROM((ScheduleRecord AS S INNER JOIN [Day] AS D ON S.DayNumber = D.DayNumber) INNER JOIN LessonTime AS L ON S.LessonTimeNumber = L.Number) INNER JOIN WeekSchedule AS WS ON S.Id = WS.ScheduleRecordId GROUP BY DayName, LessonTimePeriod, ClassRoomNumber, WS.WeekNumber HAVING COUNT(*)>1; ";

        private const string teacherInconsistenseQuery =
                "SELECT DayName,LessonTimePeriod, LastName, Initials, WS.WeekNumber AS [Week] FROM((((ScheduleRecord S INNER JOIN [Day] D ON S.DayNumber = D.DayNumber) INNER JOIN LessonTime L ON S.LessonTimeNumber = L.Number) INNER JOIN Teacher T ON S.TeacherId = T.Id) INNER JOIN WeekSchedule WS ON S.Id = WS.ScheduleRecordId) GROUP BY DayName, LessonTimePeriod, LastName, Initials, WS.WeekNumber HAVING COUNT(*)>1;";
            

        public static void ClearDataBase()
        {
            ClearTable(clearWeekScheduleQuery);
            ClearTable(clearScheduleRecordQuery);
            ClearTable(clearTeacherQuery);
        }

        public static DataTable GetInconsistentClassrooms()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = classRoomInconsistenseQuery;

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);

                return dataTable;
            }
        }

        public static DataTable GetInconsistentTeachers()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = teacherInconsistenseQuery;

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);

                return dataTable;
            }
        }

        private static void ClearTable(string query)
        {
            using (OleDbCommand oleDbCommand = new OleDbCommand())
            {
                // Set the command object properties
                oleDbCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                oleDbCommand.CommandType = CommandType.Text;
                oleDbCommand.CommandText = query;

                // Open the connection, execute the query and close the connection
                oleDbCommand.Connection.Open();
                oleDbCommand.ExecuteNonQuery();

                oleDbCommand.Connection.Close();
            }
        }


    }
}
