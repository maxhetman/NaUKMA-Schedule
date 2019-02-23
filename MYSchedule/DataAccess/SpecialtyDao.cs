using System.Data;
using System.Data.OleDb;
using System.Windows.Controls;
using MYSchedule.DTO;
using MYSchedule.Utils;

namespace MYSchedule.DataAccess
{
    public static class SpecialtyDao
    {
        private const string insertSpecialty = "Insert Into" +
                                               " Specialty(Name)" +
                                               " Values (@Name)";

        private const string getSpecialtyIdByName = "Select Id From Specialty Where Name = @Name";
        private const string getAllSpecialtiesQuery = "Select Name From Specialty";

        public static int AddIfNotExists(SpecialtyDto specialty)
        {
            var specialtyId = GetSpecialtyId(specialty);

            if (specialtyId != -1)
            {
                return specialtyId;
            }

            int result = -1;

            using (OleDbCommand oleDbCommand = new OleDbCommand())
            {
                // Set the command object properties
                oleDbCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                oleDbCommand.CommandType = CommandType.Text;
                oleDbCommand.CommandText = insertSpecialty;

                // Add the input parameters to the parameter collection
                oleDbCommand.Parameters.AddWithValue("@Name", specialty.Name);
                // Open the connection, execute the query and close the connection
                oleDbCommand.Connection.Open();
                var rowsAffected = oleDbCommand.ExecuteNonQuery();
                result = oleDbCommand.Connection.GetLatestAutonumber();
                oleDbCommand.Connection.Close();

                if (rowsAffected > 0)
                {
                    return result;
                }

                Logger.LogException("Could not add specialty");
                return -1;
            }
        }

        public static int GetSpecialtyId(SpecialtyDto specialty)
        {
            DataTable dataTable = new DataTable();
            DataRow dataRow;

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = getSpecialtyIdByName;

                // Add the parameter to the parameter collection
                dataAdapter.SelectCommand.Parameters.AddWithValue("@Name", specialty.Name);

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);

                // Get the datarow from the table
                dataRow = dataTable.Rows.Count > 0 ? dataTable.Rows[0] : null;

                return dataRow == null ? -1 : int.Parse(dataRow[0].ToString());
            }
        }

        //private bool AddSpecialty(SpecialtyDto specialty)
        //{
        //    using (OleDbCommand oleDbCommand = new OleDbCommand())
        //    {
        //        // Set the command object properties
        //        oleDbCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
        //        oleDbCommand.CommandType = CommandType.Text;
        //        oleDbCommand.CommandText = insertSpecialty;

        //        // Add the input parameters to the parameter collection
        //        oleDbCommand.Parameters.AddWithValue("@Name", specialty.Name);
        //        // Open the connection, execute the query and close the connection
        //        oleDbCommand.Connection.Open();
        //        var rowsAffected = oleDbCommand.ExecuteNonQuery();
        //        oleDbCommand.Connection.Close();

        //        return rowsAffected > 0;
        //    }
        //}
        //}
        public static string[] GetAllSpecialties()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = getAllSpecialtiesQuery;

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
