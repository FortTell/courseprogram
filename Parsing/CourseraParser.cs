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
        #endregion

        HtmlDocument document;
        public CourseraParser(StreamReader stream)
        {
            document = new HtmlDocument();
            document.Load(stream);
        }

        public ParseInfo GetAllInfo()
        {
            return new ParseInfo
            {
                courseName = GetCourseName(),
                courseDesc = GetCourseDesc(),
                teachers = GetTeachers(),
                weeks = GetWeeks()
            };
        }

        public string GetCourseName()
        {
            var name = document.DocumentNode.SelectSingleNode("/html/head/meta[@property = 'og:title']")
                .GetAttributeValue("content", "")
                .Split(" (")[0];
            return name;
        }

        public string GetCourseDesc()
        {
            var about = document.DocumentNode.SelectSingleNode(".//*[@class = 'body-1-text course-description']")
                .InnerText
                .Split(": ")[1];
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
            var npd = node.FirstChild.LastChild.InnerText
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

        public Dictionary<string, string> GetWeeks()
        {
            var weeks = document.DocumentNode.SelectNodes(".//*[@class = 'week']");
            List<string> weekInfosRaw = GetWeekInfosLinq(weeks);

            //List<string> weekInfos = GetWeekInfosSimple(weeks);
            var weekInfo = new Dictionary<string, string>();
            foreach (var wi in weekInfosRaw)
            {
                var titleMatch = titleSplitRe.Match(wi);
                weekInfo.Add(wi.Substring(0, titleMatch.Index + 1), wi.Substring(titleMatch.Index + 1));
            }
            return weekInfo;
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
