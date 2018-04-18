using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using LinqToExcel;
using MYSchedule.DTO;
using MYSchedule.Utils;
using Remotion.Mixins.CodeGeneration.DynamicProxy.TypeGeneration;

namespace MYSchedule.Parser
{
    public static class ExcelParser
    {
        private static SpecialtyDto _specialty;
        private static int _yearOfStudying;
        private static Row[] _rows;

        private static Dictionary<ScheduleRecordDto, List<int>> _weekScheduleRecords =
            new Dictionary<ScheduleRecordDto, List<int>>();

        public static Dictionary<ScheduleRecordDto, List<int>> GetScheduleFromExcel(string filePath)
        {
            var excel = new ExcelQueryFactory(filePath);
    

            var rows = from c in excel.Worksheet(0)  // getting first worksheet
                select c;         

            _rows = rows.ToArray();
            _weekScheduleRecords.Clear();
            SetSpecialtyNameAndYearOfStudying(_rows[5][0].Value.ToString());
            FillScheduleRecords(); 

            return _weekScheduleRecords;
        }

        private static void FillScheduleRecords()
        {
            var currentDay = string.Empty;
            var currentTime = string.Empty;
            DayDto currentDayDto = null;

            for (int i = 9; i < _rows.Length; i++)
            {
                var nextDay = _rows[i][0].Value.ToString();

                if (!string.IsNullOrEmpty(nextDay))
                {
                    currentDay = nextDay;
                    currentDayDto = new DayDto {DayNumber = DayDto.GetNumberByName(currentDay)};
                }

                var nextTime = _rows[i][1].Value.ToString();

                if (!string.IsNullOrEmpty(nextTime))
                {
                    currentTime = nextTime;
                }

                if (currentDay != "" && currentTime != "")
                {
                    AppendNewScheduleRecord(currentDayDto, currentTime, _rows[i]);
                }
            }
            
        }

        private static void AppendNewScheduleRecord(DayDto dayDto, string time, Row row)
        {
            try
            {
                string subject = row[2].Value.ToString();
                TeacherDto teacher = GetTeacherData(row[3].Value.ToString());
                ClassRoomDto classRoom = new ClassRoomDto { Number = row[6].Value.ToString().Replace(" ", String.Empty) };

                var weeksString = row[5].Value.ToString();

                if (string.IsNullOrEmpty(teacher.ToString()))
                {
                    teacher = new TeacherDto {LastName = "Вакансія"}; // in case of null
                }
                    

                if (string.IsNullOrEmpty(subject) || string.IsNullOrEmpty(classRoom.Number)
                    || string.IsNullOrEmpty(weeksString))
                    return;

                LessonTimeDto lessonTime = new LessonTimeDto {Number = LessonTimeDto.GetNumberFromPeriod(time)};

                int groupCheck;
                int.TryParse(row[4].Value.ToString(), out groupCheck);
                int? group = groupCheck > 0 ? (int?) groupCheck : null; //group == null if lesson type is lecture
                
                LessonTypeDto lessonType = new LessonTypeDto();
                lessonType.Id = LessonTypeDto.GetIdByType(group == null ? LessonType.Lecture : LessonType.Practice);

                ScheduleRecordDto scheduleRecord = new ScheduleRecordDto
                {
                    YearOfStudying = _yearOfStudying,
                    LessonTime = lessonTime,
                    Subject = subject,
                    LessonType = lessonType,
                    Group = group,
                    Day = dayDto,
                    Specialty = _specialty,
                    ClassRoom = classRoom,
                    Teacher = teacher,
                    Weeks = weeksString
                };

                var weeksList = Utils.Utils.ParseWeeks(row[5].Value.ToString());

                _weekScheduleRecords.Add(scheduleRecord, new List<int>());
                foreach (var weekNumber in weeksList)
                {
                   _weekScheduleRecords[scheduleRecord].Add(weekNumber);
                }

            }

            catch (Exception e)
            {
                Logger.LogException(e);
            }


        }

        private static void SetSpecialtyNameAndYearOfStudying(String columnValue)
        {
            try
            {
                var comaIndex = columnValue.IndexOf(",");
                var specialtyName = columnValue.Substring(0, comaIndex);

                _specialty = new SpecialtyDto {Name = specialtyName};

                _yearOfStudying = Convert.ToInt32(Regex.Match(columnValue.Substring(comaIndex), @"\d+").Value);
            }
            catch (Exception exc)
            {
                Logger.LogException(exc);
            }

        }

        private static void PrintOneRow(Row row)
        {
            var result = "";
            for (int i = 0; i < row.Count; i++)
            {
                result += i + ":" + row[i] + " ";

            }
            Console.WriteLine(result);
        }

        private static TeacherDto GetTeacherData(string teacherData)
        {
            var initials = "";
            var lastName = "";
            var position = "";
            teacherData = teacherData.Replace(" ", string.Empty);
            var upperCounter = 0;
            for (int i = 0; i < teacherData.Length; i++)
            {
                if (Char.IsUpper(teacherData[i]))
                {
                    if (upperCounter == 0)
                    {
                        position = teacherData.Substring(0, i);
                    }
                    if (upperCounter < 2)
                    {
                        upperCounter++;
                        initials += teacherData[i] + ".";
                    }
                    else
                    {
                        lastName = teacherData.Substring(i);
                        break;
                    }
                }
            }

            position = GetUnifiedPosition(position);

            return new TeacherDto
            {
                Initials = initials,
                LastName = lastName,
                Position = position
            };
        }

        private static string GetUnifiedPosition(string positionStr)
        {
            if (positionStr.Contains("ст") && positionStr.Contains("викл"))
            {
                return "ст. викл.";
            }

            if (positionStr.Contains("ас") || positionStr.Contains("ac"))
            {
                return "ас.";
            }

            if (positionStr.Contains("доц"))
            {
                return "доц";
            }

            return positionStr;
        }
    }


}
