namespace MYSchedule.DTO
{
    public class TeacherDto
    {
        public int Id;
        public string LastName;
        public string Initials;
        public string Position;

        public override string ToString()
        {
            return $"{nameof(Id)}: {Id}, {nameof(LastName)}: {LastName}, {nameof(Initials)}: {Initials}, {nameof(Position)}: {Position}";
        }
    }
}
