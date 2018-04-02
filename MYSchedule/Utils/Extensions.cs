using System.Data.OleDb;


namespace MYSchedule.Utils
{
    public static class Extensions
    {
        public static int GetLatestAutonumber(
            this OleDbConnection connection)
        {
            using (OleDbCommand command = new OleDbCommand("SELECT @@IDENTITY;", connection))
            {
                return (int)command.ExecuteScalar();
            }
        }
    }
}
