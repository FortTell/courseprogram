using System;
using System.Collections.Generic;
using System.Text;

namespace DataClasses
{
    public struct DisciplineInfo
    {
        public string Name;
        public int Ze;
        public int PracticeHours;
        public List<(string title, List<string> topics)> Themes;
        public bool IsExam;
    }
}
