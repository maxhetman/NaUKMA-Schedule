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

        public override string ToString()
        {
            return $"{nameof(Id)}: {Id}, {nameof(Type)}: {Type}";
        }

    }

    public enum LessonType
    {
        Lecture = 1,
        Practice = 2
    }
}
