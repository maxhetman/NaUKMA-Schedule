using System.Data;
using System.Data.OleDb;
using MYSchedule.DTO;
using MYSchedule.Utils;

namespace MYSchedule.DataAccess
{
    public static class ClassRoomsDao
    {
        private const string GetAllBuildingsQuery = "Select DISTINCT Building FROM ClassRoom";
        private const string GetAllNumbersQuery = "Select Number FROM ClassRoom";

        private const string getClassroomByNumber = "Select * From ClassRoom Where Number = @Number";

        private const string insertClassroom = "Insert Into ClassRoom([Number], NumberOfPlaces, HasProjector, IsComputerClass, Building, Board)" +
                                             " Values (@Number, 0, False, False, @Building, False)";

        public static void AddIfNotExists(ClassRoomDto classRoom)
        {
            var isExist = IsClassromExists(classRoom.Number);

            if (isExist)
                return;

            using (OleDbCommand oleDbCommand = new OleDbCommand())
            {
                // Set the command object properties
                oleDbCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                oleDbCommand.CommandType = CommandType.Text;
                oleDbCommand.CommandText = insertClassroom;

                // Add the input parameters to the parameter collection
                oleDbCommand.Parameters.AddWithValue("@Number", classRoom.Number);
                oleDbCommand.Parameters.AddWithValue("@Building", GetBuilding(classRoom.Number));
                // Open the connection, execute the query and close the connection
                oleDbCommand.Connection.Open();
                oleDbCommand.ExecuteNonQuery();
                oleDbCommand.Connection.Close();
            }
        }

        private static string GetBuilding(string number)
        {
            number = number.Replace(" ", "");
            var indexOf = number.IndexOf('-');
            return number.Substring(0, indexOf);
        }

        public static bool IsClassromExists(string classRoomNumber)
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = getClassroomByNumber;

                // Add the parameter to the parameter collection
                dataAdapter.SelectCommand.Parameters.AddWithValue("@Number", classRoomNumber);

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);

                // Get the datarow from the table
                return dataTable.Rows.Count > 0;

            }
        }

        public static string[] GetAllNumbers()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = GetAllNumbersQuery;

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);
                string[] res = new string[dataTable.Rows.Count];

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    res[i] = dataTable.Rows[i][0].ToString();
                }

                return res;
            }
        }

        public static string[] GetAllBuildings()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = GetAllBuildingsQuery;

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);
                string[] res = new string[dataTable.Rows.Count];

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    res[i] = dataTable.Rows[i][0].ToString();
                }

                return res;
            }
        }
    }
}
