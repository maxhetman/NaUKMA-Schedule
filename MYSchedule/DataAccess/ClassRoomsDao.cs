using System.Data;
using System.Data.OleDb;

namespace MYSchedule.DataAccess
{
    public static class ClassRoomsDao
    {
        private const string GetAllBuildingsQuery = "Select DISTINCT Building FROM ClassRoom";
        private const string GetAllNumbersQuery = "Select Number FROM ClassRoom";

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
