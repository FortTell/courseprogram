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
        public static DocX BuildDocx(string url)
        {
            using (var r = new StreamReader(HttpWebRequest.Create(url).GetResponse().GetResponseStream()))
                return BuildDocx(r);
        }
        public static DocX BuildDocx(StreamReader reader)
        {
            var parser = new CourseraParser(reader);
            var courseName = parser.GetCourseName();

            if (!Directory.Exists("\\out"))
                Directory.CreateDirectory("\\out");

            using (var d = DocX.Create("out\\" + courseName + ".docx"))
            {
                var nameP = d.InsertParagraph()
                    .FontSize(14).Bold().Font(new Font("Arial"));
                nameP.Alignment = Alignment.center;
                nameP.Append(courseName + "\n\n" + "123");
                
                //var teachers = parser.GetTeachers();
                d.Save();
            }
            throw new NotImplementedException();
        }
    }
}