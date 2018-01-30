using System;
using System.Collections.Generic;
using System.Text;

namespace Parsing
{
    public struct TeacherInfo
    {
        public String name, degree, position, department;

        public override string ToString()
        {
            return String.Join(", ", name, degree, position, department);
        }
    }
}
