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
                SetTableBorder(teacherT);

                d.Save();
            }
        }

        private static void SetTableBorder(Table t)
        {
            foreach (TableBorderType borderType in Enum.GetValues(typeof(TableBorderType)))
                t.SetBorder(borderType,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
        }

        private static void FillEducResults(DocX template)
        {
            var rnd = new Random();
            var educResTable = template.Tables[3];
            using (var d = DocX.Load("competence.docx"))
            {
                var t = d.Tables[0];
                var rowCount = t.RowCount;
                for (int i = 0; i < (rowCount / 2) + 1; i++)
                {
                    educResTable.InsertRow();
                    var rowIndex = rnd.Next(0, t.RowCount);
                    var row = t.Rows[rowIndex];
                    educResTable.Rows[educResTable.RowCount - 1].Cells[1].Paragraphs[0]
                        .Append(row.Cells[0].Paragraphs[0].Text);
                    var parasCount = row.Cells[1].Paragraphs.Count;
                    var cell = educResTable.Rows[educResTable.RowCount - 1].Cells[2];
                    cell.RemoveParagraphAt(0);
                    for (int j = 0; j < (parasCount / 2) + 1; j++)
                    {
                        var para = row.Cells[1].Paragraphs[rnd.Next(0, row.Cells[1].Paragraphs.Count)];
                        cell.InsertParagraph(para);
                        row.Cells[1].RemoveParagraph(para);
                    }
                    t.RemoveRow(rowIndex);
                }
            }
            SetTableBorder(educResTable);
        }

        public static void BuildDocxFromTemplate(ParseInfo pi, string templateFilename)
        {
            if (!Directory.Exists("out"))
                Directory.CreateDirectory("out");
            LoadReplacements(pi);
                

            using (var d = DocX.Load(templateFilename))
            {
                d.InsertDocument(DocX.Load("disc_template.docx"));
                FillEducResults(d);
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
                { "<DISC_NAME>", pi.courseName }, //no support for multi-discipline modules yet
                { "<YEAR>", DateTime.Now.Year.ToString()},
                { "<CITY>", "Екатеринбург" },
                { "<NAME>", "Иванов И.И." },
                { "<POSITION>", "Искатель интересных историй" },
                { "<UNIVERSITY_NAME>", "ВУЗ им. Иванова И.И." },
                { "<SEMESTER_NO>", "5"},
                { "<TEST_TYPE>", "экзамен"}
            };
        }
    }
}