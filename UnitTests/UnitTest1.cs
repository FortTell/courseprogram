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
        TeacherInfo teacher1 = new TeacherInfo { name = "����", degree = "�������� ����", department = "������", position = "���.��������" };
        TeacherInfo noDegree = new TeacherInfo { name = "�������", position = "�������" };
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
            Assert.IsTrue(name == "����������� �������������");
        }

        [TestMethod]
        public void SimpleTeacherParsing()
        {
            var teachers = mainParser.GetTeachers();
            Assert.IsTrue(teachers[0].name == "������ ������������");
            Assert.IsTrue(teachers[0].department.Contains("����"));
            Assert.IsNull(teachers[1].degree);
            Assert.IsTrue(teachers[1].department.Contains("����"));
        }

        [TestMethod]
        public void TeacherToStringIgnoresEmptyFields()
        {
            Assert.IsFalse(noDegree.ToString().Contains(", ,"));
            Assert.IsTrue(noDegree.ToString() == "�������, �������");
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
