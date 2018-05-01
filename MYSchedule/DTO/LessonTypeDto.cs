using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MYSchedule.DTO
{
    public class LessonTypeDto
    {
        public int Id; 
        public string Type; //lecture/practice

        public static int GetIdByType(LessonType type)
        {
            return (int) type;
        }

        public override int GetHashCode()
        {
            return (Type != null ? Type.GetHashCode() : 0);
        }

    }

    public enum LessonType
    {
        Лекція = 1,
        Практика = 2
    }
}
