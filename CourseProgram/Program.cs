using Parsing;
using System;
using System.IO;
using DocxBuilder;
using System.Collections.Generic;
using DataClasses;
using System.Diagnostics;
using System.Net;
using System.Text.RegularExpressions;
using System.Reflection;
using CommandLine;

namespace CourseProgram
{
    public class Program
    {
        static void Main(string[] args)
        {
            var gHandler = InitProgram(args);
            Console.WriteLine("Info from webpage parsed\nLink: " + gHandler.SheetLink);
            var p = new Process();
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = gHandler.SheetLink;
            p.Start();
            Console.WriteLine("Please fill the necessary information at Link: \n Then press Enter to create a document, other to abort");
            var key = Console.ReadKey();
            if (!(key.Key == ConsoleKey.Enter))
                return;
            Console.WriteLine("Creating document...");
            Builder.BuildDocxFromTemplate(gHandler.Parser.ParseInfoFromSheet(), "template.docx");
        }

        private static GSheetApiHandler InitProgram(string[] args)
        {
            Console.WriteLine("Program started.");
            Directory.SetCurrentDirectory(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location));

            string handlerStr = "";
            bool isURL = false;
            var cliOpt = Parser.Default.ParseArguments<CliOptions>(args);
            try
            {
                cliOpt.MapResult<CliOptions, int>(cli =>
                {
                    if (!(cli.Filename == null ^ cli.WebLink == null))
                        throw new ArgumentException("Please specify exactly one of the -f, -l params.");
                    handlerStr = cli.Filename ?? cli.WebLink;
                    isURL = cli.Filename == null;
                    return 0;
                }, _ => throw new ArgumentException("Cannot parse command line args to work, exiting."));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadKey();
                Environment.Exit(0);
            }
            return PrepareHandlerAndMakeFirstPass(handlerStr, isURL);
        }

        private static GSheetApiHandler PrepareHandlerAndMakeFirstPass(string source, bool sourceIsWeblink)
        {
            if (!sourceIsWeblink)
                return MakeFirstParsePass(File.OpenRead(source));
            else
            {
                var webCl = new WebClient();
                Stream resp = webCl.OpenRead(source);
                return MakeFirstParsePass(resp);
            }
        }

        public static GSheetApiHandler MakeFirstParsePass(Stream reader)
        {
            var parser = new CourseraParser(reader);
            var pi = parser.ParseInfoFromWebpage();
            var gHandler = new GSheetApiHandler();
            gHandler.InitParser();
            gHandler.PasteInfoToSheet(pi, 0);
            return gHandler;
        }

        public static void PrintHelp()
        {

        }
    }
}
