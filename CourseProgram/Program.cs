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
            //GSheetParser.ConnectToSheet();
            //Console.ReadLine();
            MakeDocx(File.OpenRead("course.html"));
        }

        public static void MakeDocx(Stream reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();
            var pi = parser.ParseInfoFromWebpage();
            var gParser = new GSheetParser();
            gParser.PasteParseInfoToSheet(pi, 0);
            Builder.BuildDocxFromTemplate(gParser.ParseInfoFromSheet(), "template.docx");
        }
    }
}
