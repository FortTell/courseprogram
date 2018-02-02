using Parsing;
using System;
using System.IO;
using DocxBuilder;
using System.Collections.Generic;
using DataClasses;

namespace CourseProgram
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //BuildDocx(new StreamReader("tests\\course.html"));
        }

        public static void MakeDocx(StreamReader reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();
            var pi = new ParseInfo { courseName = courseName, teachers = parser.GetTeachers() };
            Builder.BuildDocx(pi);
        }
    }
}
