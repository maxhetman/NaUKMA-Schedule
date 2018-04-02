using System;
using System.IO;

namespace MYSchedule.Utils
{
    public static class Logger
    {

        private static string DirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "Schedule");

        private static string FilePath = Path.Combine(DirectoryPath, "Error.txt");

        public static void LogException(Exception exc)
        {
            if (!Directory.Exists(DirectoryPath))
                Directory.CreateDirectory(DirectoryPath);
            using (StreamWriter writer = new StreamWriter(FilePath, true))
            {
                writer.WriteLine("Message :" + exc.Message + "<br/>" + Environment.NewLine + "StackTrace :" + exc.StackTrace +
                                 "" + Environment.NewLine + "Date :" + DateTime.Now);
                writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
            }
        }

        public static void LogException(string exc)
        {
            if (!Directory.Exists(DirectoryPath))
                Directory.CreateDirectory(DirectoryPath);

            using (StreamWriter writer = new StreamWriter(FilePath, true))
            {
                writer.WriteLine("Message :" + exc + "<br/>" + Environment.NewLine
                    + Environment.NewLine + "Date :" + DateTime.Now);
                writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
            }
        }

    }
}
