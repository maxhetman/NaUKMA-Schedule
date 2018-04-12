namespace MYSchedule.DataAccess
{
    public static class Queries
    {
        private const string classroomsBusyness = @"Select D.DayName, L.LessonTimePeriod, S.ClassRoomNumber, T.LastName, T.Initials, S.Weeks From ((((ScheduleRecord S Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join ClassRoom C ON S.ClassRoomNumber = C.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) ";

        private const string OrderByLessonTimeClassRoom = " ORDER BY L.Number, C.Number";



        private const string scheduleForLessonByCourseSpecialtySubject =
              "Select [W.Number] AS WeekNumber, D.DayName, L.Number, S.ClassRoomNumber, T.LastName, T.Initials, LT.Type, S.Group From(((((((ScheduleRecord S Inner Join WeekSchedule WS ON S.Id = WS.ScheduleRecordId) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join LessonType LT ON S.LessonTypeId = LT.Id) Inner Join [Specialty] SP ON S.SpecialtyId = SP.Id) Inner Join [Week] W ON WS.WeekNumber = W.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) ";

        private const string scheduleForWeek =
                "SELECT DayName AS День, LessonTimePeriod AS Пара,  Subject AS Предмет, LT.Type AS Тип, SP.Name AS Спеціальність, YearOfStudying AS Курс, Group AS Група, ClassRoomNumber AS Аудиторія FROM(((((ScheduleRecord AS S INNER JOIN [Day] AS D ON S.DayNumber= D.DayNumber) INNER JOIN LessonTime AS L ON S.LessonTimeNumber=L.Number) INNER JOIN Teacher AS T ON S.TeacherId = T.Id) INNER JOIN Specialty AS SP ON S.SpecialtyId=SP.Id) INNER JOIN WeekSchedule AS WS ON S.Id = WS.ScheduleRecordId) INNER JOIN LessonType AS LT ON S.LessonTypeId = LT.Id ";

        private const string OrderByDayLessonClassRoomWeek = " ORDER BY D.DayNumber, L.Number, S.ClassRoomNumber, W.Number";
        private const string OrderByDayLessonSubject = " ORDER BY D.DayNumber, L.Number, Subject";

        public static string ScheduleForWeekQuery(int weekNumber)
        {
            var query = scheduleForWeek;
            query += "WHERE WeekNumber = " + weekNumber + OrderByDayLessonSubject;
            return query;
        }

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
                query += "Where C.Number = \"" + classroomNumber + "\"";
            }

            query += OrderByLessonTimeClassRoom;
            return query;
        }
    }
}