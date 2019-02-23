using System;

namespace MYSchedule.Exceptions
{
    public class InvalidInputException: Exception
    {
        public InvalidInputException(string message)
            : base(message)
        {
        }

    }
}
