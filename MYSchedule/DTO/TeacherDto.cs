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

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (LastName != null ? LastName.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Initials != null ? Initials.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Position != null ? Position.GetHashCode() : 0);
                return hashCode;
            }
        }
    }
}
