using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;
using System.IO;
using HtmlAgilityPack;
using System.Linq;

namespace CourseProgram
{
    public class Parser
    {
        HtmlDocument document = new HtmlDocument();
        Regex weekRe = new Regex(@"WEEK \d+");
        Regex weekContentRe = new Regex(@"\d+ videos.*");
        Regex titleRe = new Regex("[а-я][А-Я]");

        public Parser(StreamReader stream)
        {
            document.Load(stream);
        }

        public Dictionary<string, string> GetWeeks()
        {
            var weeks = document.DocumentNode.SelectNodes(@".//*[@class = 'week']");
            List<string> weekInfos = GetWeekInfosLinq(weeks);

            //List<string> weekInfos = GetWeekInfosSimple(weeks);
            var dict = new Dictionary<string, string>();
            foreach (var wi in weekInfos)
            {
                var titleMatch = titleRe.Match(wi);
                dict.Add(wi.Substring(0, titleMatch.Index + 1), wi.Substring(titleMatch.Index + 1));
            }
            return dict;
        }

        private List<string> GetWeekInfosSimple(HtmlNodeCollection weeks)
        {
            var weekInfos = new List<string>();
            foreach (var node in weeks)
            {
                Console.WriteLine(node.InnerText.Split("expand")[0]);
                var info = node.InnerText.Split("expand")[0];
                var weekMatch = weekRe.Match(info);
                var weekContentMatch = weekContentRe.Match(info);
                if (weekContentMatch.Success)
                    weekInfos.Add(info.Substring(weekMatch.Index + weekMatch.Length, weekContentMatch.Index - (weekMatch.Index + weekMatch.Length)));
                else
                    weekInfos.Add(info.Substring(weekMatch.Index + weekMatch.Length));
            }
            return weekInfos;
        }

        private List<string> GetWeekInfosLinq(HtmlNodeCollection weeks)
        {
            return weeks
                .Select(w => w.InnerText.Split("expand")[0])
                .Select(w => w.Substring(weekRe.Match(w).Index + weekRe.Match(w).Length))
                .Select(w => weekContentRe.Match(w).Success ? w.Substring(0, weekContentRe.Match(w).Index) : w)
                .ToList();
        }
    }
}
