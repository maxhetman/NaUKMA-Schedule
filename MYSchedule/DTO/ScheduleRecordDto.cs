using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MYSchedule.DTO
{
    public class ScheduleRecordDto
    {
        public int Id;
        public int YearOfStudying; //1-6
        public LessonTimeDto LessonTime; //FK LessonDto    

        public string Subject;
        public LessonTypeDto LessonType; //FK LessonTypeDto
        public string Group; //null if lecture 

        public DayDto Day;
        public SpecialtyDto Specialty;
        public ClassRoomDto ClassRoom;
        public TeacherDto Teacher;
        public string Weeks;

        public override string ToString()
        {
            return
                $"{nameof(Id)}: {Id}," +
                $" {nameof(YearOfStudying)}: {YearOfStudying}, {nameof(LessonTime)}: {LessonTime}, " +
                $"{nameof(Subject)}: {Subject}, {nameof(LessonType)}: {LessonType}, {nameof(Group)}: {Group}," +
                $" {nameof(Day)}: {Day}, {nameof(Specialty)}: {Specialty}, {nameof(ClassRoom)}: {ClassRoom}," +
                $" {nameof(Teacher)}: {Teacher}, {nameof(Weeks)}: {Weeks}";
        }

        protected bool Equals(ScheduleRecordDto other)
        {
            return Id == other.Id && YearOfStudying == other.YearOfStudying && Equals(LessonTime, other.LessonTime) && string.Equals(Subject, other.Subject) && Equals(LessonType, other.LessonType) && Group == other.Group && Equals(Day, other.Day) && Equals(Specialty, other.Specialty) && Equals(ClassRoom, other.ClassRoom) && Equals(Teacher, other.Teacher);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((ScheduleRecordDto) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = 7;
                hashCode = (hashCode * 397) ^ YearOfStudying;
                hashCode = (hashCode * 397) ^ (LessonTime != null ? LessonTime.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Subject != null ? Subject.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (LessonType != null ? LessonType.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Group != null ? Group.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Day != null ? Day.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Specialty != null ? Specialty.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (ClassRoom != null ? ClassRoom.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Teacher != null ? Teacher.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Weeks != null ? Weeks.GetHashCode() : 0);
                return hashCode;
            }
        }

    }
}
