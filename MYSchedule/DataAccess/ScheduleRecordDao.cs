using System;
using System.Data;
using System.Data.OleDb;
using MYSchedule.DTO;
using MYSchedule.Utils;

namespace MYSchedule.DataAccess
{
    public static class ScheduleRecordDao
    {
        private const string insertScheduleRecord = "Insert Into" +
                                                    " ScheduleRecord(Id, YearOfStudying, [Subject], LessonTypeId, [Group], " +
                                                    "TeacherId, DayNumber, ClassRoomNumber, LessonTimeNumber, SpecialtyId, Weeks)" +
                                                    " Values (@Id, @YearOfStudying, @Subject, @LessonTypeId, @Group, @TeacherId," +
                                                    " @DayNumber, @ClassRoomNumber, @LessonTimeNumber, @SpecialtyId, @Weeks)";

        private const string selectScheduleRecordById = "Select * From ScheduleRecord Where Id = @Id";

        private const string getAllSpecialtiesQuery = "Select DISTINCT Subject FROM ScheduleRecord";
        private const string getAllYearsQuery = "Select DISTINCT YearOfStudying FROM ScheduleRecord";
        public static bool AddIfNotExists(ScheduleRecordDto scheduleRecord)
        {

            var specialtyId = SpecialtyDao.AddIfNotExists(scheduleRecord.Specialty);
            var teacherId = TeacherDao.AddIfNotExists(scheduleRecord.Teacher);
            ClassRoomsDao.AddIfNotExists(scheduleRecord.ClassRoom);
            scheduleRecord.Id = scheduleRecord.GetHashCode();

            if (IsStoredInDb(scheduleRecord.Id))
            {
                return false;
            }
            try
            {
                using (OleDbCommand oleDbCommand = new OleDbCommand())
                {
                    // Set the command object properties
                    oleDbCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                    oleDbCommand.CommandType = CommandType.Text;
                    oleDbCommand.CommandText = insertScheduleRecord;

                    // Add the input parameters to the parameter collection
                    oleDbCommand.Parameters.AddWithValue("@Id", scheduleRecord.Id);
                    oleDbCommand.Parameters.AddWithValue("@YearOfStudying", scheduleRecord.YearOfStudying);
                    oleDbCommand.Parameters.AddWithValue("@Subject", scheduleRecord.Subject);
                    oleDbCommand.Parameters.AddWithValue("@LessonTypeId", scheduleRecord.LessonType.Id);
                    if (scheduleRecord.Group == string.Empty)
                    {
                        oleDbCommand.Parameters.AddWithValue("@Group", DBNull.Value);
                    }
                    else
                    {
                        oleDbCommand.Parameters.AddWithValue("@Group", scheduleRecord.Group);
                    }
                    oleDbCommand.Parameters.AddWithValue("@TeacherId", teacherId);
                    oleDbCommand.Parameters.AddWithValue("@DayNumber", scheduleRecord.Day.DayNumber);
                    oleDbCommand.Parameters.AddWithValue("@ClassRoomNumber", scheduleRecord.ClassRoom.Number);
                    oleDbCommand.Parameters.AddWithValue("@LessonTimeNumber", scheduleRecord.LessonTime.Number);
                    oleDbCommand.Parameters.AddWithValue("@SpecialtyId", specialtyId);
                    oleDbCommand.Parameters.AddWithValue("@Weeks", scheduleRecord.Weeks);

                    // Open the connection, execute the query and close the connection
                    oleDbCommand.Connection.Open();
                    var rowsAffected = oleDbCommand.ExecuteNonQuery();

                    oleDbCommand.Connection.Close();

                    if (rowsAffected > 0)
                    {
                        return true;
                    }

                    Logger.LogException("Could not add schedule record");
                    return false;
                }
            }
            catch (OleDbException ex)
            {
                Logger.LogException(ex);
                return false;
            }
        }

        public static string[] GetAllSubjects()
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

        public static string[] GetAllYears()
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = getAllYearsQuery;

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
        public static bool IsStoredInDb(int scheduleRecordId)
        {
            DataTable dataTable = new DataTable();

            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter())
            {
                // Create the command and set its properties
                dataAdapter.SelectCommand = new OleDbCommand();
                dataAdapter.SelectCommand.Connection = new OleDbConnection(ConnectionConfig.ConnectionString);
                dataAdapter.SelectCommand.CommandType = CommandType.Text;
                dataAdapter.SelectCommand.CommandText = selectScheduleRecordById;

                // Add the parameter to the parameter collection
                dataAdapter.SelectCommand.Parameters.AddWithValue("@Id", scheduleRecordId);

                // Fill the datatable From adapter
                dataAdapter.Fill(dataTable);

                // Get the datarow from the table
                return dataTable.Rows.Count > 0;
            }
        }
    }

}
