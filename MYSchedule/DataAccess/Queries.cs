namespace MYSchedule.DataAccess
{
    public static class Queries
    {
        private const string classroomsBusyness = @"Select D.DayName, L.LessonTimePeriod, S.ClassRoomNumber, T.LastName, T.Initials, S.Weeks From ((((ScheduleRecord S Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join ClassRoom C ON S.ClassRoomNumber = C.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) ";

        private const string OrderByLessonTimeClassRoom = " ORDER BY L.Number, C.Number";



        private const string scheduleForLessonByCourseSpecialtySubject =
                "Select [W.Number] AS WeekNumber, D.DayName, L.LessonTimePeriod, S.ClassRoomNumber, T.LastName, T.Initials, LT.Type, S.Group From(((((((ScheduleRecord S Inner Join WeekSchedule WS ON S.Id = WS.ScheduleRecordId) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join LessonType LT ON S.LessonTypeId = LT.Id) Inner Join [Specialty] SP ON S.SpecialtyId = SP.Id) Inner Join [Week] W ON WS.WeekNumber = W.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) "
            ;

        private const string OrderByDayLessonClassRoomWeek = " ORDER BY D.DayNumber, L.Number, S.ClassRoomNumber, W.Number";

        public static string LessonScheduleByCourseSpecialtySubjectQuery(string specialty, int yearOfStudying, string subject)
        {
            var query = scheduleForLessonByCourseSpecialtySubject;
            query += string.Format("Where SP.Name={0} And S.YearOfStudying={1} And S.Subject={2}", specialty,
                yearOfStudying, subject);
            query += OrderByDayLessonClassRoomWeek;
            return query;
        }
            public static string ClassRoomsAvailabilityQuery(int? buildingNumber, bool? isComputer,
            string classroomNumber)
        {
            var query = classroomsBusyness;

            if (buildingNumber != null)
            {
                query += "Where Building = " + buildingNumber;
            }

            if (isComputer != null)
            {
                query += buildingNumber == null
                    ? "Where IsComputerClass = " + isComputer
                    : " And IsComputerClass = " + isComputer;
            }

            if (classroomNumber != null)
            {
                query += "Where Number = " + classroomNumber;
            }

            query += OrderByLessonTimeClassRoom;
            return query;
        }
    }
}