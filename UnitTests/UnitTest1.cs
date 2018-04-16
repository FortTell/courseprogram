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
            mainParser = new CourseraParser(File.OpenRead("tests\\course.html"));
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
        public void NameDoesNotEndInASemicolon()
        {
            var parser = new CourseraParser(File.OpenRead("tests\\endingSemicolon.html"));
            var teachers = parser.GetTeachers();
            foreach (var t in teachers)
                Assert.IsTrue(t.ToString().Last() != ',');
        }

        [TestMethod]
        public void BracketsInNamesAreHandled()
        {
            var parser = new CourseraParser(File.OpenRead("tests\\bracketedNames.html"));
            var teachers = parser.GetTeachers();
            foreach (var t in teachers)
            {
                Assert.IsTrue(!t.name.Contains(')'));
                Assert.IsTrue(!t.name.Contains('('));
            };
        }

        [TestMethod]
        public void PerformanceTest()
        {
            Program.MakeFirstParsePass(File.OpenRead("tests\\course.html"));

        }

        [TestMethod]
        public void CheckHourSpreadingForOutliers()
        {
            var themes = new List<(string, List<string>)>();
            for (int i = 0; i < 5; i++)
                themes.Add(("t" + i, new List<string> { "st" + i }));
            var d = DisciplineInfo.CreateSecondPassDI("test", 3, (20, 40, 30), themes, true);
            
        }
    }
}
