using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;

namespace MYSchedule.DTO
{
    public class ClassRoomDto
    {
        public string Number;//1-102, 3-304
        public int NumberOfPlaces;
        public bool HasProjector;
        public bool IsComputerClass;
        public int Building; //depends on Number (1, 3)

        public override string ToString()
        {
            return $"{nameof(Number)}: {Number}";
        }
    }
}
