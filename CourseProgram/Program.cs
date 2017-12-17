using System;
using System.IO;

namespace CourseProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            /*var parser = new CourseraParser(new System.IO.StreamReader("tests\\course.html"));
            var weeks = parser.GetWeeks();
            var about = parser.GetCourseDesc();
            var teachers = parser.GetTeachers();
            var courseName = parser.GetCourseName();*/
            DocxBuilder.BuildDocx(new StreamReader("tests\\course.html"));
            Console.WriteLine("Hello World!");
        }
    }
}
