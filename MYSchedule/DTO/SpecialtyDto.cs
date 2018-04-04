using System.Collections.Generic;

namespace MYSchedule.DTO
{
    public class SpecialtyDto
    {
        public int Id; 
        public string Name;

        public override string ToString()
        {
            return $"{nameof(Id)}: {Id}, {nameof(Name)}: {Name}";
        }

        public override int GetHashCode()
        {
            return (Name != null ? Name.GetHashCode() : 0);
        }
    }
}
