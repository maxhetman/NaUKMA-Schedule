namespace MYSchedule.DataAccess
{
    public static class Queries
    {
        private const string classroomsAvailabilityForAllWeeks = @"Select D.DayName AS День, L.LessonTimePeriod AS Пара, S.ClassRoomNumber AS Аудиторія, T.LastName, T.Initials, S.Weeks AS Тижні From ((((ScheduleRecord S Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join ClassRoom C ON S.ClassRoomNumber = C.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) ";

        private const string classroomsAvailabilityForAllWeeks2 = @"SELECT dayLessonClassRoom.DayName AS День, dayLessonClassRoom.LessonTimePeriod AS Пара, dayLessonClassRoom.ClassRoomNumber AS Аудиторія, Teacher.LastName AS Прізвище, Teacher.Initials AS Ініціали, ScheduleRecord.Weeks AS Тижні FROM ( SELECT ClassRoom.Number AS ClassRoomNumber, ClassRoom.Building, Day.DayNumber, Day.DayName, LessonTime.Number AS LessonTimeNumber, LessonTime.LessonTimePeriod FROM ClassRoom, [Day], LessonTime) AS dayLessonClassRoom LEFT JOIN (ScheduleRecord LEFT JOIN Teacher ON ScheduleRecord.TeacherId = Teacher.Id) ON (dayLessonClassRoom.DayNumber = ScheduleRecord.DayNumber) AND (dayLessonClassRoom.LessonTimeNumber = ScheduleRecord.LessonTimeNumber) AND (dayLessonClassRoom.ClassRoomNumber = ScheduleRecord.ClassRoomNumber) ";

        private const string classroomsAvailabilityForSelectedWeek2 = @"SELECT dayLessonClassRoom.DayName AS День, dayLessonClassRoom.LessonTimePeriod AS Пара, dayLessonClassRoom.ClassRoomNumber AS Аудиторія, Teacher.LastName AS Прізвище, Teacher.Initials AS Ініціали, WS.WeekNumber AS Тиждень FROM ( SELECT ClassRoom.Number AS ClassRoomNumber, ClassRoom.Building, Day.DayNumber, Day.DayName, LessonTime.Number AS LessonTimeNumber, LessonTime.LessonTimePeriod FROM ClassRoom, [Day], LessonTime) AS dayLessonClassRoom LEFT JOIN ((ScheduleRecord LEFT JOIN Teacher ON ScheduleRecord.TeacherId = Teacher.Id) LEFT JOIN WeekSchedule WS ON ScheduleRecord.Id = WS.ScheduleRecordID) ON (dayLessonClassRoom.DayNumber = ScheduleRecord.DayNumber) AND (dayLessonClassRoom.LessonTimeNumber = ScheduleRecord.LessonTimeNumber) AND (dayLessonClassRoom.ClassRoomNumber = ScheduleRecord.ClassRoomNumber) ";

        private const string classroomsAvailabilityForSelectedWeek = @"Select D.DayName AS День, L.LessonTimePeriod AS Пара, S.ClassRoomNumber AS Аудиторія, T.LastName, T.Initials, WS.WeekNumber AS Тиждень From (((((ScheduleRecord S Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join ClassRoom C ON S.ClassRoomNumber = C.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) Inner Join WeekSchedule WS ON S.Id = WS.ScheduleRecordId) ";

        private const string OrderByLessonTimeClassRoom = " ORDER BY D.DayNumber, L.Number, C.Number";

        private const string scheduleForLessonByCourseSpecialtySubject =
              "Select [W.Number] AS WeekNumber, D.DayName, L.Number, S.ClassRoomNumber, T.LastName, T.Initials, LT.Type, S.Group From(((((((ScheduleRecord S Inner Join WeekSchedule WS ON S.Id = WS.ScheduleRecordId) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join LessonType LT ON S.LessonTypeId = LT.Id) Inner Join [Specialty] SP ON S.SpecialtyId = SP.Id) Inner Join [Week] W ON WS.WeekNumber = W.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) ";

        private const string scheduleForWeek =
                "SELECT DayName AS День, LessonTimePeriod AS Пара,  Subject AS Предмет, LT.Type AS Тип, SP.Name AS Спеціальність, YearOfStudying AS Курс, Group AS Група, ClassRoomNumber AS Аудиторія FROM(((((ScheduleRecord AS S INNER JOIN [Day] AS D ON S.DayNumber= D.DayNumber) INNER JOIN LessonTime AS L ON S.LessonTimeNumber=L.Number) INNER JOIN Teacher AS T ON S.TeacherId = T.Id) INNER JOIN Specialty AS SP ON S.SpecialtyId=SP.Id) INNER JOIN WeekSchedule AS WS ON S.Id = WS.ScheduleRecordId) INNER JOIN LessonType AS LT ON S.LessonTypeId = LT.Id ";

        private const string OrderByDayLessonClassRoomWeek = " ORDER BY D.DayNumber, L.Number, S.ClassRoomNumber, W.Number";

        private const string OrderByDayLessonSubject = " ORDER BY D.DayNumber, L.Number, Subject";

        private const string teacherScheduleForAllWeeks = "SELECT DayName AS День, LessonTimePeriod AS Пара, ClassRoomNumber AS Аудиторія, Subject AS Предмет, LT.Type AS Тип, SP.Name AS Спеціальність, YearOfStudying AS Курс, Group AS Група,  Weeks AS Тижні FROM((((ScheduleRecord AS S INNER JOIN [Day] AS D ON S.DayNumber= D.DayNumber) INNER JOIN LessonTime AS L ON S.LessonTimeNumber=L.Number) INNER JOIN Teacher AS T ON S.TeacherId = T.Id) INNER JOIN Specialty AS SP ON S.SpecialtyId=SP.Id) INNER JOIN LessonType AS LT ON S.LessonTypeId=LT.Id ";

        private const string teacherScheduleForSelectedWeek =
                "SELECT DayName AS День, LessonTimePeriod AS Пара, ClassRoomNumber AS Аудиторія, Subject AS Предмет, LT.Type AS Тип, SP.Name AS Спеціальність, YearOfStudying AS Курс, Group AS Група FROM(((((ScheduleRecord AS S INNER JOIN [Day] AS D ON S.DayNumber= D.DayNumber) INNER JOIN LessonTime AS L ON S.LessonTimeNumber=L.Number) INNER JOIN Teacher AS T ON S.TeacherId = T.Id) INNER JOIN Specialty AS SP ON S.SpecialtyId=SP.Id) INNER JOIN WeekSchedule AS WS ON S.Id = WS.ScheduleRecordId) INNER JOIN LessonType AS LT ON S.LessonTypeId = LT.Id "
            ;

        private const string studentScheduleForAllWeeks =
                "SELECT DayName AS День,  LessonTimePeriod AS Пара, ClassRoomNumber AS Аудиторія, T.LastName & \" \" & T.Initials AS Вчитель , Subject AS Предмет, LT.Type AS Тип, Group AS Група, Weeks AS Тижні FROM((((ScheduleRecord AS S INNER JOIN [Day] AS D ON S.DayNumber= D.DayNumber) INNER JOIN LessonTime AS L ON S.LessonTimeNumber=L.Number) INNER JOIN Teacher AS T ON S.TeacherId = T.Id) INNER JOIN LessonType AS LT ON S.LessonTypeId=LT.Id) INNER JOIN Specialty SP ON S.SpecialtyId = SP.Id "
            ;

        private const string studentScheduleForSelectedWeek =
                "SELECT DayName AS День,  LessonTimePeriod AS Пара, ClassRoomNumber AS Аудиторія,T.LastName & \" \" & T.Initials AS Вчитель , Subject AS Предмет, LT.Type AS Тип, Group AS Група FROM(((((ScheduleRecord AS S INNER JOIN [Day] AS D ON S.DayNumber= D.DayNumber) INNER JOIN LessonTime AS L ON S.LessonTimeNumber=L.Number) INNER JOIN Teacher AS T ON S.TeacherId = T.Id) INNER JOIN LessonType AS LT ON S.LessonTypeId=LT.Id) INNER JOIN WeekSchedule WS ON S.Id = WS.ScheduleRecordId) INNER JOIN Specialty SP ON S.SpecialtyId = SP.Id ";

        public static string ScheduleForWeekQuery(int weekNumber)
        {
            var query = scheduleForWeek;
            query += "WHERE WeekNumber = " + weekNumber + OrderByDayLessonSubject;
            return query;
        }

        public static string TeacherScheduleForAllWeeksQuery(string teacher, string initials)
        {
            var query = teacherScheduleForAllWeeks;
            query += "WHERE T.[LastName] = \"" + teacher + "\" AND T.[Initials] = \"" + initials + "\"" + OrderByDayLessonSubject;
            return query;
        }

        public static string StudentScheduleForAllWeeksQuery(string specialtyName, int course)
        {
            var query = studentScheduleForAllWeeks;
            query += "WHERE SP.Name = \"" + specialtyName + "\" AND YearOfStudying = " + course + OrderByDayLessonSubject;
            return query;
        }

        public static string TeacherScheduleForSelectedWeekQuery(string teacher, string initials, int weekNumber)
        {
            var query = teacherScheduleForSelectedWeek;
            query += "WHERE T.[LastName] = \"" + teacher + "\" AND T.[Initials] = \"" + initials + "\" AND WeekNumber = "+ weekNumber + OrderByDayLessonSubject;
             return query;
        }

        public static string OrderByMethodistFirstQuery =
                " ORDER BY dayLessonClassRoom.DayNumber, dayLessonClassRoom.LessonTimeNumber, dayLessonClassRoom.Building, dayLessonClassRoom.ClassRoomNumber"
            ;
        public static string StudentScheduleForSelectedWeekQuery(string specialtyName, int weekNumber, int course)
        {
            var query = studentScheduleForSelectedWeek;
            query += "WHERE WS.WeekNumber = " + weekNumber + " AND YearOfStudying = " + course + " AND SP.Name = \"" + specialtyName + "\"";
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


        public static string ClassRoomsAvailabilityForAllWeeksQuery(int? buildingNumber, bool? isComputer,
            string classroomNumber)
        {
            var query = classroomsAvailabilityForAllWeeks2;

            if (buildingNumber != null)
            {
                query += "Where dayLessonClassRoom.Building = " + buildingNumber;
            }

            if (isComputer != null)
            {
                query += buildingNumber == null
                    ? "Where dayLessonClassRoom.IsComputerClass = " + isComputer
                    : " And dayLessonClassRoom.IsComputerClass = " + isComputer;
            }

            if (classroomNumber != null)
            {
                query += "Where dayLessonClassRoom.ClassRoomNumber = \"" + classroomNumber + "\"";
            }

            query += OrderByMethodistFirstQuery;
            return query;
        }

        public static string ClassRoomsAvailabilityForSelectedWeekQuery(int? buildingNumber, bool? isComputer,
            string classroomNumber, int week)
        {
            var query = classroomsAvailabilityForSelectedWeek2;
            var filterPartOfQuery = string.Empty;

            if (buildingNumber != null)
            {
                filterPartOfQuery += "Where dayLessonClassRoom.Building = " + buildingNumber;
            }

            if (isComputer != null)
            {
                filterPartOfQuery += buildingNumber == null
                    ? "Where dayLessonClassRoom.IsComputerClass = " + isComputer
                    : " And dayLessonClassRoom.IsComputerClass = " + isComputer;
            }

            if (classroomNumber != null)
            {
                filterPartOfQuery += "Where dayLessonClassRoom.ClassRoomNumber = \"" + classroomNumber + "\"";
            }

            filterPartOfQuery = filterPartOfQuery == string.Empty
                ? "Where WS.WeekNumber = " + week
                : " And WS.WeekNumber = " + week;

            query += filterPartOfQuery;
            query += OrderByMethodistFirstQuery;
            return query;
        }
    }
}