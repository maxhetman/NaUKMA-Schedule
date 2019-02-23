using System;
using System.Collections.Generic;
using System.Linq;
using MYSchedule.Utils;

namespace MYSchedule.DTO
{
    public class LessonTimeDto
    {
        public int Number; //1,2,..,7
        public string LessonTimePeriod; //8:30-9:50...



        private static Dictionary<int, string> LessonTimeToNumber;

        #region Utils


        static LessonTimeDto()
        {
            LessonTimeToNumber = new Dictionary<int, string>()
            {
                {1, "8:30-9:50"},
                {2, "10:00-11:20"},
                {3, "11:40-13:00"},
                {4, "13:30-14:50"},
                {5, "15:00-16:20"},
                {6, "16:30-17:50"},
                {7, "18:00-19:20"}
            };
        }

        public static int GetNumberFromPeriod(string period)
        {
            period = period.Replace(" ", String.Empty).Replace(".", ":");

            foreach (KeyValuePair<int, string> entry in LessonTimeToNumber)
            {
                if (entry.Value == period)
                    return entry.Key;
            }

            Logger.LogException("[LessonDto] Wrong LessonTimePeriod: " + period);
            return -1;

;        }

        public static string GetPeriodFromNumber(int number)
        {
            return LessonTimeToNumber[number];
        }

        #endregion    

        public override int GetHashCode()
        {
            return (LessonTimePeriod != null ? LessonTimePeriod.GetHashCode() : 0);
        }

    }
}
