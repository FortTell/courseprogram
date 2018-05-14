using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;

namespace CourseProgram
{
    public class CliOptions
    {
        [Option('f', "filename", HelpText = "Downloaded course page location")]
        public string Filename { get; set; }

        [Option('l', "url", HelpText = "Course page URL")]
        public string WebLink { get; set; }

    }
}
