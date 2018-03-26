using System;
using System.Collections.Generic;
using System.Text;

namespace DataClasses
{
    public struct DisciplineInfo
    {
        public string Name { get; private set; }
        public int Ze { get; private set; }
        public int TotalHours => Ze * 36;
        public (int practice, int seminar, int selfWork) Hours { get; private set; }
        public List<(string title, List<string> topics)> Themes { get; private set; }
        public bool IsExam { get; private set; }

        public static DisciplineInfo CreateFirstPassDI(string name, List<(string, List<string>)> themes)
        {
            return new DisciplineInfo { Name = name, Themes = themes };
        }

        public static DisciplineInfo CreateSecondPassDI(string name, int ze, (int,int,int) hours, 
            List<(string, List<string>)> themes, bool isExam)
        {
            return new DisciplineInfo
            {
                Name = name,
                Ze = ze,
                Hours = hours,
                Themes = themes,
                IsExam = isExam
            };
        }
    }
}
