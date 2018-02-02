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
        public Dictionary<string, string> weeks;
    }
}
