using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataClasses
{
    public struct TeacherInfo
    {
        public String name, degree, position, department;

        public override string ToString()
        {
            return String.Join(", ", 
                new[] { name, degree, position, department }
                    .Where(x => x != "" && x != null)
            );
        }
    }
}
