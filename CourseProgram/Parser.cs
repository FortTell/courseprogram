using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using HtmlAgilityPack;

namespace CourseProgram
{
    public class Parser
    {
        HtmlDocument document = new HtmlDocument();
        public Parser(StreamReader stream)
        {
            document.Load(stream);
        }

        public List<string> GetWeeks()
        {
            var weeks = document.DocumentNode.SelectNodes(@".//*[@class = 'week']");
            foreach (var node in weeks)
            {
                
                Console.WriteLine(node.InnerText.Split("expand")[0]);
            }

            throw new NotImplementedException();
        }
    }
}
