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

            if (DT == null) return new DataTable();

            return DT;
        }

        public static DataTable GetScheduleBySubjectSpecialtyAndCourse(string specialty, int course, string subject)
        {
            string query = Queries.LessonScheduleByCourseSpecialtySubjectQuery(specialty, course, subject);

            DataTable DT = Access2Dt(query);

            if (DT == null) return new DataTable();

            return DT;
        }

        public static DataTable GetScheduleForWeek(int week)
        {
            string query = Queries.ScheduleForWeekQuery(week);

            DataTable DT = Access2Dt(query);

            if (DT == null) return new DataTable();

            return DT;
        }

        public static DataTable GetTeacherScheduleForAllWeeks(string teacherLastName, string initials)
        {
            string query = Queries.TeacherScheduleForAllWeeksQuery(teacherLastName, initials);

            DataTable DT = Access2Dt(query);

            if (DT == null) return new DataTable();

            return DT;
        }

        public static DataTable GetStudentScheduleForAllWeeks(string specialtyName)
        {
            string query = Queries.StudentScheduleForAllWeeksQuery(specialtyName);

            DataTable DT = Access2Dt(query);

            if (DT == null) return new DataTable();

            return DT;
        }

        public static DataTable GetTeacherScheduleForSelectedWeek(string teacherLastName, string initials, int weekNumber)
        {
            string query = Queries.TeacherScheduleForSelectedWeekQuery(teacherLastName, initials, weekNumber);

            DataTable DT = Access2Dt(query);

            if (DT == null) return new DataTable();

            return DT;
        }

        public static DataTable GetStudentScheduleForSelectedWeek(string specialty, int week)
        {
            string query = Queries.StudentScheduleForSelectedWeekQuery(specialty, week);

            DataTable DT = Access2Dt(query);

            if (DT == null) return new DataTable();

            return DT;
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