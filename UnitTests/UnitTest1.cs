using Microsoft.VisualStudio.TestTools.UnitTesting;
using Parsing;
using DocxBuilder;
using CourseProgram;
using DataClasses;
using System.Collections.Generic;

namespace UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        TeacherInfo t1 = new TeacherInfo { name = "Иван", degree = "кандидат наук", department = "Лестех", position = "зав.кафедрой" };

        [TestMethod]
        public void TestMethod1()
        {
            var pi = new ParseInfo { courseName = "abc", teachers = new List<TeacherInfo> { t1 } };
            
        }
    }
}
