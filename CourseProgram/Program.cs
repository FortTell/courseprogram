using Parsing;
using System;
using System.IO;
using DocxBuilder;
using System.Collections.Generic;
using DataClasses;
using System.Diagnostics;

namespace CourseProgram
{
    public class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");
            var gParser = MakeFirstParsePass(File.OpenRead("course.html"));
            Console.WriteLine("Info from webpage parsed\nLink: " + gParser.SheetLink);
            //Process.Start(gParser.SheetLink);
            Console.WriteLine("Press Enter to continue, other to abort");
            var key = Console.ReadKey();
            if (!(key.Key == ConsoleKey.Enter))
                return;
            Console.WriteLine("Creating document...");
            Builder.BuildDocxFromTemplate(gParser.ParseInfoFromSheet(), "template.docx");
        }

        public static GSheetParser MakeFirstParsePass(Stream reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();
            var pi = parser.ParseInfoFromWebpage();
            var gParser = new GSheetParser();
            gParser.PasteInfoToSheet(pi, 0);
            return gParser;
        }
    }
}
