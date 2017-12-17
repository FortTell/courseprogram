using System;

namespace CourseProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            var parser = new CourseraParser(new System.IO.StreamReader("course.html"));
            var weeks = parser.GetWeeks();
            var about = parser.GetCourseDesc();
            var teachers = parser.GetTeachers();
            
            Console.WriteLine("Hello World!");
        }
    }
}
