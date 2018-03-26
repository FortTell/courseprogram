using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataClasses
{
    public struct ParseInfo
    {
        public string CourseName;
        public string CourseDesc;
        public List<TeacherInfo> Teachers;
        public List<DisciplineInfo> Disciplines;
        public int ModuleZe { get => Disciplines.Select(d => d.Ze).Sum(); }
    }
}
