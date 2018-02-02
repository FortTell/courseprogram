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
            MakeDocx(new StreamReader("course.html"));
        }

        public static void MakeDocx(StreamReader reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();
            var pi = parser.ParseAllInfo();
            //            Builder.BuildDocx(pi);
            Builder.BuildDocxFromTemplate(pi);
        }
    }
}
