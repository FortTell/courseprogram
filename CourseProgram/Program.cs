using Parsing;
using System;
using System.IO;
using DocxBuilder;
using System.Collections.Generic;

namespace CourseProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            BuildDocx(new StreamReader("tests\\course.html"));
        }

        public static void BuildDocx(StreamReader reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();
            var pi = new ParseInfo { courseName = courseName, teachers = new List<string> { "", "" } };
            DocxBuilder.DocxBuilder.BuildDocx(pi);
        }
    }
}
