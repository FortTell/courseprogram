using System;
using System.Collections.Generic;
using System.Text;

namespace DataClasses
{
    public struct ParseInfo
    {
        public string courseName;
        public string courseDesc;
        public List<TeacherInfo> teachers;
        public List<(string title, List<string> topics)> themes;
        public int moduleZe;
    }
}
