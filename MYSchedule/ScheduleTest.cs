using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using MYSchedule.DataAccess;
using MYSchedule.DTO;
using MYSchedule.ExcelExport;
using MYSchedule.Parser;

namespace MYSchedule
{
    class ScheduleTest
    {
        public static void Main(string[] args)
        {


            //Stopwatch stopWatch = new Stopwatch();
            //Console.WriteLine("Start parsing");
            //stopWatch.Start();
            //var schedule = ExcelParser.GetScheduleFromExcel();
            //stopWatch.Stop();
            //Console.WriteLine("Milis passed : " + stopWatch.Elapsed.Milliseconds);
            //var count = 0;

            //foreach (KeyValuePair<ScheduleRecordDto, List<int>> entry in schedule)
            //{
            //    stopWatch.Restart();
            //    bool isAdded = ScheduleRecordDao.AddIfNotExists(entry.Key);

            //    if (!isAdded)
            //    {
            //        continue;
            //    }

            //    foreach (var weekNumber in entry.Value)
            //    {
            //        WeekScheduleDao.AddWeekSchedule(weekNumber: weekNumber, scheduleRecordId: entry.Key.Id);
            //    }
            //    stopWatch.Stop();
            //    count++;
            //    Console.WriteLine("Milis passed : " + stopWatch.Elapsed.Milliseconds + " for " + count + " record");
            //}

            Console.OutputEncoding = Encoding.UTF8;
            var dt = QueryManager.GetClassRoomsAvailability();

            foreach (DataRow one in dt.Rows)
            {
                Console.WriteLine($"{one[0]}, {one[1]}, {one[2]}, {one[3]}, {one[4]}, {one[5]}");
            }

            ExcelExportManager.ShowAllClassRooms(dt);

            Console.ReadLine();

            //fileName = @"E:\kek.mdb";
            //var kek = DataBaseCreator.CreateNewAccessDatabase(fileName);
            //new SpecialtyDao().AddIfNotExists(new SpecialtyDto
            //{
            //    Name = "lol"
            //});

            //var a = new SpecialtyDao().GetSpecialtyId(new SpecialtyDto
            //{
            //    Name = "Kek"
            //});
            //var a = TeacherDao.AddIfNotExists(new TeacherDto {Initials = "a.m", LastName = "GLYB", Position = "BOSS"});            
            //Console.WriteLine(a);
            //Console.WriteLine(TeacherDao.GetTeacherId(new TeacherDto{Initials = "a.m", LastName = "GLYB"}));
            //Console.ReadLine();
        }

        private static void WriteElapsed(TimeSpan ts)
        {
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            Console.WriteLine();
        }
    }
}
