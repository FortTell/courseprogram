using Microsoft.VisualStudio.TestTools.UnitTesting;
using Parsing;
using DocxBuilder;
using CourseProgram;
using DataClasses;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        TeacherInfo teacher1 = new TeacherInfo { name = "Иван", degree = "кандидат наук", department = "Лестех", position = "зав.кафедрой" };
        TeacherInfo noDegree = new TeacherInfo { name = "Василий", position = "уборщик" };
        CourseraParser mainParser;

        [TestInitialize]
        public void ParserInit()
        {
            mainParser = new CourseraParser(new StreamReader("tests\\course.html"));
        }

        [TestMethod]
        public void CourseNameOK()
        {
            var name = mainParser.GetCourseName();
            Assert.IsTrue(name == "Современная комбинаторика");
        }

        [TestMethod]
        public void SimpleTeacherParsing()
        {
            var teachers = mainParser.GetTeachers();
            Assert.IsTrue(teachers[0].name == "Андрей Райгородский");
            Assert.IsTrue(teachers[0].department.Contains("МФТИ"));
            Assert.IsNull(teachers[1].degree);
            Assert.IsTrue(teachers[1].department.Contains("МФТИ"));
        }

        [TestMethod]
        public void TeacherToStringIgnoresEmptyFields()
        {
            Assert.IsFalse(noDegree.ToString().Contains(", ,"));
            Assert.IsTrue(noDegree.ToString() == "Василий, уборщик");
        }

        [TestMethod]
        public void NameEndingWithSemicolonWorks()
        {
            var parser = new CourseraParser(new StreamReader("tests\\endingSemicolon.html"));
            var teachers = parser.GetTeachers();
            foreach (var t in teachers)
                Assert.IsTrue(t.ToString().Last() != ',');
        }
    }
}
