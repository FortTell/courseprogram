using Novacode;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace CourseProgram
{
    public static class DocxBuilder
    {
        public static void BuildDocx(string url)
        {
            using (var r = new StreamReader(HttpWebRequest.Create(url).GetResponse().GetResponseStream()))
                BuildDocx(r);
        }
        public static void BuildDocx(StreamReader reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();

            if (!Directory.Exists("out"))
                Directory.CreateDirectory("out");

            using (var d = DocX.Create("out\\" + courseName + ".docx"))
            {
                var nameP = d.InsertParagraph();
                nameP.Alignment = Alignment.center;
                nameP.Append(courseName + new String('\n', 5))
                    .FontSize(14).Bold().Font(new Font("Arial"));
                
                var teachers = parser.GetTeachers();
                d.Save();
            }
        }
    }
}