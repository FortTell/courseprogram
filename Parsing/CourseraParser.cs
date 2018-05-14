using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DataClasses;

namespace Parsing
{
    public class CourseraParser
    {
        #region Regexes
        Regex weekRe = new Regex(@"WEEK \d+");
        Regex weekContentRe = new Regex(@"\d+ videos.*");
        Regex titleSplitRe = new Regex("[а-я][А-Я]");
        Regex degreeRe = new Regex(@"([Дд]октор)|([Кк]андидат).*наук");
        Regex attestRe = new Regex(@"([Зз]ач[её]т)|([Ээ]кзамен)");
        Regex videoTopicRe = new Regex(@"ideo: .*?[^ (][A-Z]");
        #endregion

        HtmlDocument document;
        public CourseraParser(Stream stream)
        {
            document = new HtmlDocument();
            document.Load(stream);
        }

        public ParseInfo ParseInfoFromWebpage()
        {
            return new ParseInfo
            {
                CourseName = GetCourseName(),
                CourseDesc = GetCourseDesc(),
                Teachers = GetTeachers(),
                Disciplines = new List<DisciplineInfo> {
                    DisciplineInfo.CreateFirstPassDI(GetCourseName(), GetThemes())
                }
            };
        }

        public string GetCourseName()
        {
            var name = document.DocumentNode.SelectSingleNode("/html/head/meta[@property = 'og:title']")
                .GetAttributeValue("content", "")
                .Split(new string[] { " (", " |" }, StringSplitOptions.None)[0];
            return name;
        }

        public string GetCourseDesc()
        {
            var about = document.DocumentNode
                .SelectSingleNode(".//*[@class = 'about-section-wrapper']" +
                "//*[@class = 'body-1-text course-description']/text()")
                .InnerText;
            return about;
        }

        public List<TeacherInfo> GetTeachers()
        {
            var teachers = document.DocumentNode.SelectNodes(".//*[contains(@class,'instructor-info')]");
            var teacherInfo = ParseTeacherInfo(teachers);
            return teacherInfo;
        }

        private List<TeacherInfo> ParseTeacherInfo(HtmlNodeCollection teacherNodes)
        {
            var result = new List<TeacherInfo>();
            foreach (var node in teacherNodes)
            {
                var ti = ParseNamePositionAndDegree(node);
                ti.department = node.LastChild.InnerText;
                result.Add(ti);
            }
            return result;
        }

        private TeacherInfo ParseNamePositionAndDegree(HtmlNode node)
        {
            var npd = Regex.Replace(node.FirstChild.LastChild.InnerText, "[(].*?[)]", "")
                                .Replace("&nbsp;", "")
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => x.TrimStart())
                                .ToList<string>();
            var ti = new TeacherInfo { name = npd[0] };
            if (npd.Count == 1)
                return ti;
            for (int i = 1; i < npd.Count; i++)
            {
                if (degreeRe.Match(npd[i]).Success)
                    ti.degree += (npd[i].ToLower() + (ti.degree == null ? "" : ", "));
                else
                    ti.position += (npd[i] + (ti.position == null ? "" : ", "));
            }
            return ti;
        }

        public List<(string title, List<string> topics)> GetThemes()
        {
            var weeks = document.DocumentNode.SelectNodes(".//*[@class = 'week']");
            var weekInfosRaw = GetWeekInfosLinq(weeks);
            var leftToCompress = weekInfosRaw.Count - 5;
            var totalTopicLength = weekInfosRaw.Select(k => k.topics.Sum(t => t.Length)).Sum();
            for (int i = 0; leftToCompress > 0; i++)
                while (leftToCompress > 0 && weekInfosRaw[i].topics.Sum(t => t.Length) < totalTopicLength / 5)
                {
                    leftToCompress--;
                    weekInfosRaw[i] = (weekInfosRaw[i].title + ". " + weekInfosRaw[i + 1].title,
                        weekInfosRaw[i].topics.Concat(weekInfosRaw[i + 1].topics).ToList());
                    weekInfosRaw.RemoveAt(i + 1);
                }
            return weekInfosRaw;
        }

        private List<(string title, List<string> topics)> GetWeekInfosLinq(HtmlNodeCollection weeks)
        {
            var antiDict = File.ReadAllLines("AntiDictionary.txt");
            return weeks
                .SelectMany(w => w.LastChild.ChildNodes)
                .Select(w => (w.ChildNodes[0].InnerText,
                    videoTopicRe.Matches(w.InnerText)
                        .Select(m => m.Value.Substring(6, m.Value.Length - 7))
                        .Where(s => !antiDict.Any(antiWord => s.IndexOf(antiWord, StringComparison.OrdinalIgnoreCase) >= 0))
                        .ToList()))
                .Where(tup => !(attestRe.IsMatch(tup.Item1))).ToList();
            /*return weeks
                .Select(w => w.LastChild.LastChild)
                .Select(w => (w.ChildNodes[0].InnerText,
                    w.ChildNodes[1].InnerText.Split("More")[0]
                    .Replace('\t', ' ').Split(". ").ToList()))
                .Where(tup => !(attestRe.IsMatch(tup.Item1))).ToList()
                .ToList();*/
        }
    }
}
