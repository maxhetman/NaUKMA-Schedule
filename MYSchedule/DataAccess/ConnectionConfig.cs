using System.Configuration;

namespace MYSchedule.DataAccess
{
    public static class ConnectionConfig
    {
        public static string ConnectionString
        {
            get
            {
                return ConfigurationManager
                    .ConnectionStrings["ScheduleDBConnection"]
                    .ToString();
            }
        }

    }
}
