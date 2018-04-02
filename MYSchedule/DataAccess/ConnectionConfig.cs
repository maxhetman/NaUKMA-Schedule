using System.Configuration;

namespace MYSchedule.DataAccess
{
    public static class ConnectionConfig
    {
        public static string ConnectionString => ConfigurationManager
            .ConnectionStrings["ScheduleDBConnection"]
            .ToString();
    }
}
