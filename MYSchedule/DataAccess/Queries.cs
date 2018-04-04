namespace MYSchedule.DataAccess
{
    public static class Queries
    {
        private const string classroomsBusyness = @"Select D.DayName, L.LessonTimePeriod, S.ClassRoomNumber, T.LastName, T.Initials, S.Weeks From ((((ScheduleRecord S Inner Join [Day] D ON S.DayNumber = D.DayNumber) Inner Join Teacher T ON S.TeacherId = T.Id) Inner Join ClassRoom C ON S.ClassRoomNumber = C.Number) Inner Join LessonTime L ON S.LessonTimeNumber = L.Number) ";

        private const string OrderByLessonTimeClassRoom = " ORDER BY L.Number, C.Number";

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