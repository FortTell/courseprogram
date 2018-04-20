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
            var gHandler = MakeFirstParsePass(File.OpenRead("course.html"));
            Console.WriteLine("Info from webpage parsed\nLink: " + gHandler.SheetLink);
            var p = new Process();
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = gHandler.SheetLink;
            p.Start();
            Console.WriteLine("Press Enter to continue, other to abort");
            var key = Console.ReadKey();
            if (!(key.Key == ConsoleKey.Enter))
                return;
            Console.WriteLine("Creating document...");
            Builder.BuildDocxFromTemplate(gHandler.Parser.ParseInfoFromSheet(), "template.docx");
        }

        public static GSheetApiHandler MakeFirstParsePass(Stream reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();
            var pi = parser.ParseInfoFromWebpage();
            var gHandler = new GSheetApiHandler();
            gHandler.InitParser();
            gHandler.Parser.PrepareInfoToPasteToSheet(pi, 0);
            return gHandler;
        }
    }
}
