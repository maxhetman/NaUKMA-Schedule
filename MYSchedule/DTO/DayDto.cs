using System;
using MYSchedule.Utils;

namespace MYSchedule.DTO
{
    public class DayDto
    {
        public int DayNumber; //0-6
        public string DayName; //Monday-Sunday

        public static int GetNumberByName(string name)
        {
            name = name.Replace(" ", String.Empty);

            switch (name)
            {
                case Constants.Monday:
                    return 1;
                case Constants.Tuesday:
                    return 2;
                case Constants.Wednesday:
                    return 3;
                case Constants.Thursday:
                    return 4;
                case Constants.Friday:
                    return 5;
                case Constants.Saturday:
                    return 6;
                default:
                    if (name.EndsWith("тниця"))
                    {
                        return 5;
                    }
                    else
                    {
                        Logger.LogException("Could not found day number for: " + name);
                        return -1;
                    }
            }           
        }
        public override string ToString()
        {
            return $"{nameof(DayNumber)}: {DayNumber}, {nameof(DayName)}: {DayName}";
        }

        public override int GetHashCode()
        {
            return (DayName != null ? DayName.GetHashCode() : 0);
        }
    }
}
