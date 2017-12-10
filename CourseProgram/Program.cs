using System;

namespace CourseProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            var parser = new Parser(new System.IO.StreamReader("course.html"));
            parser.GetWeeks();
            Console.WriteLine("Hello World!");
        }
    }
}
