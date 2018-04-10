using System;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using MYSchedule.Utils;

namespace MYSchedule.DataAccess
{
    public static class QueryManager
    {

        public static DataTable GetClassRoomsAvailability(int? buildingNumber = null, bool? isComputer = null,
            string classroomNumber = null)
        {
            string query = Queries.ClassRoomsAvailabilityQuery(buildingNumber, isComputer, classroomNumber) ;
            
            DataTable DT = Access2Dt(query);

            if (DT == null || DT.Rows.Count < 1) throw new Exception("Data not found");
            else return DT;
        }

        public static DataTable GetScheduleBySubjectSpecialtyAndCourse(string specialty, int course, string subject)
        {
            string query = Queries.LessonScheduleByCourseSpecialtySubjectQuery(specialty, course, subject);
            DataTable DT = Access2Dt(query);
            if (DT == null || DT.Rows.Count < 1) throw new Exception("Data not found");
            else return DT;
        }

        private static DataTable Access2Dt(string query)
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                oleDbDataAdapter.SelectCommand = new OleDbCommand();
                oleDbDataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                oleDbDataAdapter.SelectCommand.CommandType = CommandType.Text;

                // Assign the SQL to the command object
                oleDbDataAdapter.SelectCommand.CommandText = query;

                // Fill the datatable from adapter
                oleDbDataAdapter.Fill(dataTable);
            }

            return dataTable;
        }
    }
}