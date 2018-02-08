using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using DataClasses;
using System.Linq;

namespace DocxBuilder
{
    public static class Builder
    {
        public static Dictionary<string, string> TemplateReplacements { get; private set; }

        public static void BuildDocx(ParseInfo pi)
        {
            if (!Directory.Exists("out"))
                Directory.CreateDirectory("out");
            using (var d = DocX.Create("out\\" + pi.courseName + ".docx"))
            {
                var nameP = d.InsertParagraph();
                nameP.Alignment = Alignment.center;
                nameP.Append(pi.courseName + new String('\n', 5))
                    .FontSize(14).Bold().Font(new Novacode.Font("Arial"));

                var teachers = pi.teachers;
                var teacherT = d.InsertTable(2, 6);
                foreach (TableBorderType borderType in Enum.GetValues(typeof(TableBorderType)))
                    teacherT.SetBorder(borderType,
                        new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));

                d.Save();
            }
        }

        public static void BuildDocxFromTemplate(ParseInfo pi, string templateFilename)
        {
            if (!Directory.Exists("out"))
                Directory.CreateDirectory("out");
            LoadReplacements(pi);

            using (var d = DocX.Load(templateFilename))
            {
                var paragraphs = d.Paragraphs.Where(p => p.Text != "");
                foreach (var p in paragraphs)
                    foreach (var repl in TemplateReplacements)
                        p.ReplaceText(repl.Key, repl.Value);
                d.SaveAs("out\\" + pi.courseName + ".docx");
            }
        }

        private static void LoadReplacements(ParseInfo pi)
        {

            TemplateReplacements = new Dictionary<string, string>
            {
                { "<COURSE_NAME>", pi.courseName },
                { "<YEAR>", DateTime.Now.Year.ToString()},
                { "<CITY>", "Екатеринбург" },
                { "<NAME>", "Иванов И.И." },
                { "<POSITION>", "Искатель интересных историй" },
                { "<UNIVERSITY_NAME>", "ВУЗ им. Иванова И.И." }
            };
        }
    }
}