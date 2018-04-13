using System.Data;
using System.Data.OleDb;
using MYSchedule.DTO;
using MYSchedule.Utils;

namespace MYSchedule.DataAccess
{
    public class TeacherDao
    {
        private const string insertTeacher = "Insert Into Teacher(LastName, Initials, [Position])" +
                                             " Values (@LastName, @Initials, @Position)";

        private const string getTeacher = "Select Id From Teacher Where LastName = @LastName AND Initials = @Initials";

        private const string getAllTeachers = "Select LastName, Initials From Teacher ORDER BY LastName";

        public static int AddIfNotExists(TeacherDto teacher)
        {
            var teacherId = GetTeacherId(teacher);

            if (teacherId != -1)
            {
                return teacherId;
            }

            int result = -1;

            using (OleDbCommand oleDbCommand = new OleDbCommand())
            {
                // Set the command object properties
                oleDbCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                oleDbCommand.CommandType = CommandType.Text;
                oleDbCommand.CommandText = insertTeacher;

                // Add the input parameters to the parameter collection
                oleDbCommand.Parameters.AddWithValue("@LastName", teacher.LastName);
                oleDbCommand.Parameters.AddWithValue("@Initials", teacher.Initials);
                oleDbCommand.Parameters.AddWithValue("@Position", teacher.Position);
                // Open the connection, execute the query and close the connection
                oleDbCommand.Connection.Open();
                var rowsAffected = oleDbCommand.ExecuteNonQuery();
                result = oleDbCommand.Connection.GetLatestAutonumber();
                oleDbCommand.Connection.Close();

                if (rowsAffected > 0)
                {
                    return result;
                }

                Logger.LogException("Could not add teacher");
                return -1;
            }
        }

        public static string[] GetFormattedTeachers()
        {
            var teachersDt = GetAllTeachers();
            string[] res = new string[teachersDt.Rows.Count];
            for (int i = 0; i < res.Length; i++)
            {
                res[i] = GetFormattedTeacherInfo(teachersDt.Rows[i]);
            }
            return res;
        }

        private static string GetFormattedTeacherInfo(DataRow teachersDtRow)
        {
            return teachersDtRow[0] + " " + teachersDtRow[1];
        }

        public static DataTable GetAllTeachers()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = getAllTeachers;

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);
                return dataTable;
            }
        }

        public static int GetTeacherId(TeacherDto teacher)
        {
            DataTable dataTable = new DataTable();
            DataRow dataRow;

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = getTeacher;

                // Add the parameter to the parameter collection
                dataAdapter.SelectCommand.Parameters.AddWithValue("@LastName", teacher.LastName);
                dataAdapter.SelectCommand.Parameters.AddWithValue("@Initials", teacher.Initials);

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);

                // Get the datarow from the table
                dataRow = dataTable.Rows.Count > 0 ? dataTable.Rows[0] : null;

                return dataRow == null ? -1 : int.Parse(dataRow[0].ToString());
            }
        }
    }
}
