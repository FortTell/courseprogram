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
        public static int TablesInDiscTempl = 10;

        private static void FillEducResults(DocX template, int discId)
        {
            FillModuleEducResTable(template.Tables[3 + TablesInDiscTempl * discId]);
            FillCompetByDiscTable(template, discId);
        }

        private static void FillModuleEducResTable(Table educResTable)
        {
            var rnd = new Random();
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
                    var parasToAdd = new List<Paragraph>();
                    for (int j = 0; j < (parasCount / 2) + 1; j++)
                    {
                        var para = row.Cells[1].Paragraphs[rnd.Next(0, row.Cells[1].Paragraphs.Count)];
                        parasToAdd.Add(para);
                        row.Cells[1].RemoveParagraph(para);
                    }
                    parasToAdd = parasToAdd
                        .OrderBy(p => p.Text.Split('-')[0])
                        .ThenBy(p => int.Parse(p.Text.Split(new char[] { '-', ' ' })[1])).ToList();
                    for (int j = 0; j < parasToAdd.Count; j++)
                        cell.InsertParagraph(parasToAdd[j]);
                    t.RemoveRow(rowIndex);
                }
                SetTableBorder(educResTable);
            }
        }
        private static void FillCompetByDiscTable(DocX template, int discId)
        {
            var erTable = template.Tables[3 + TablesInDiscTempl * discId];
            var cbdTable = template.Tables[4 + TablesInDiscTempl * discId];
            var competences = erTable.Rows.Skip(1).SelectMany(r => r.Cells[2].Paragraphs)
                .OrderBy(p => p.Text.Split('-')[0])
                .ThenBy(p => int.Parse(p.Text.Split(new char[] { '-', ' ' })[1]))
                .ToList();
            var competCodes = competences.Select(c => c.Text.Split()[0]).ToList();
            for (int i = 0; i < Math.Min(8, competCodes.Count); i++)
            {
                cbdTable.InsertColumn();
                cbdTable.Rows[0].Cells[1 + i].Paragraphs[0].Append(competCodes[i]).Bold();
                cbdTable.Rows[1].Cells[2 + i].Paragraphs[0].Append("*").Bold().Alignment = Alignment.center;
                template.Lists[4].Items[2].AppendLine(competences[i].Text).FontSize(12); //*///
            }
            SetTableBorder(cbdTable);
        }
        private static void FillDiscContentTable(DocX template, ParseInfo pi, int discId)
        {
            var rnd = new Random();
            var disc = pi.Disciplines[discId];
            var dcTable = template.Tables[9 + TablesInDiscTempl * discId];
            for (int i = 0; i < 5; i++)
            {
                dcTable.InsertRow();
                var row = dcTable.Rows[i + 1];
                row.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                row.Cells[0].Paragraphs[0].Append((i + 1).ToString());
                var para = row.Cells[1].Paragraphs[0];
                para.Append(disc.Themes[i].title + (disc.Themes[i].title.EndsWith('.') ? "" : "."));
                var topics = disc.Themes[i].topics.ToList();
                for (int k = 0; k < Math.Max(topics.Count / 4.5, 2); k++)
                {
                    var topic = topics[rnd.Next(0, topics.Count)];
                    row.Cells[2].Paragraphs[0].Append(topic + (topic.EndsWith('.') ? " " : ". "));
                    topics.Remove(topic);
                }
            }
        }

        private static void FillHourTables(DocX template, ParseInfo pi, int discId)
        {
            var d = pi.Disciplines[discId];
            var hours = SpreadHours(d);
            FillDiscSizeTable(template, d, discId);
            FillHourSpreadTable(template, d, discId, hours);
        }
        private static (int, int, int)[] SpreadHours(DisciplineInfo d)
        {
            var rnd = new Random();
            var hourShares = new int[d.Themes.Count];
            var hours = new(int, int, int)[d.Themes.Count];
            for (int i = 0; i < d.Themes.Count - 1; i++)
            {
                var percentLeft = (100 - hourShares.Sum()) / (d.Themes.Count - i);
                hourShares[i] = rnd.Next(percentLeft - 7, percentLeft + 7);
                hours[i] = (hourShares[i] * d.Hours.practice / 100,
                    hourShares[i] * d.Hours.seminar / 100,
                    hourShares[i] * d.Hours.selfWork / 100);
            }
            hourShares[d.Themes.Count - 1] = 100 - hourShares.Sum();
            hours[d.Themes.Count - 1] = (d.Hours.practice - hours.Sum(h => h.Item1),
                d.Hours.seminar - hours.Sum(h => h.Item2),
                d.Hours.selfWork - hours.Sum(h => h.Item3));
            return hours;
        }
        private static void FillDiscSizeTable(DocX template, DisciplineInfo d, int discId)
        {
            var table = template.Tables[8 + TablesInDiscTempl * discId];
            var contactHrs = 0d;
            AppendToCell(table, 2, 2, d.Hours.practice);
            AppendToCell(table, 2, 3, d.Hours.practice);
            contactHrs += d.Hours.practice;
            AppendToCell(table, 2, 4, d.Hours.practice);
            AppendToCell(table, 4, 2, d.Hours.practice);
            AppendToCell(table, 4, 3, d.Hours.practice);
            contactHrs += d.Hours.practice;
            AppendToCell(table, 4, 4, d.Hours.practice);
            AppendToCell(table, 6, 2, d.Hours.seminar + d.Hours.selfWork);
            AppendToCell(table, 6, 3, Math.Round((d.Hours.seminar + d.Hours.selfWork) * 0.15d).ToString());
            contactHrs += Math.Round((d.Hours.seminar + d.Hours.selfWork) * 0.15d);
            AppendToCell(table, 6, 4, d.Hours.seminar + d.Hours.selfWork);
            AppendToCell(table, 7, 2, d.AttestHours);
            AppendToCell(table, 7, 3, d.AttestHours);
            contactHrs += d.AttestHours;
            AppendToCell(table, 7, 4, (d.IsExam ? "Э(" : "З(") + d.AttestHours + ")");
            AppendToCell(table, 8, 2, d.TotalHours);
            AppendToCell(table, 8, 3, contactHrs.ToString());
            AppendToCell(table, 8, 4, d.TotalHours);
        }
        private static void FillHourSpreadTable(DocX template, DisciplineInfo d, int discId,
            (int pra, int sem, int swk)[] hours)
        {
            var table = template.Tables[10 + TablesInDiscTempl * discId];

            PrepareHourSpreadTable(table, d);
            var spentHours = new int[3];
            for (int i = 0; i < d.Themes.Count; i++)
            {
                var row = table.Rows[4 + i];
                AppendToCell(table, 4 + i, 0, i + 1, 8);
                AppendToCell(table, 4 + i, 1, d.Themes[i].title, 8);
                AppendToCell(table, 4 + i, 3, hours[i].pra, 8);
                AppendToCell(table, 4 + i, 5, hours[i].pra, 8);
                AppendToCell(table, 4 + i, 8, hours[i].sem, 8);
                AppendToCell(table, 4 + i, 10, hours[i].sem, 8);
                AppendToCell(table, 4 + i, 13, hours[i].swk, 8);
                AppendToCell(table, 4 + i, 14, hours[i].swk, 8);
            }
        }
        private static void PrepareHourSpreadTable(Table table, DisciplineInfo d)
        {
            for (int i = 1; i < d.Themes.Count; i++)
                table.InsertRow(table.Rows[4], 4 + i);
            for (int i = 27; i <= 30; i++)
                table.MergeCellsInColumn(i, 3, 3 + 1 + d.Themes.Count);

            AppendToCell(table, table.RowCount - 1, 2, d.TotalHours.ToString(), 8);
            AppendToCell(table, table.RowCount - 2, 2, (d.TotalHours - (d.IsExam ? 18 : 4)).ToString(), 8);
            AppendToCell(table, table.RowCount - 1, 3, d.Hours.practice.ToString(), 8);
            AppendToCell(table, table.RowCount - 2, 3, d.Hours.practice.ToString(), 8);
            AppendToCell(table, table.RowCount - 1, 5, (d.Hours.selfWork + d.Hours.seminar).ToString(), 8);
            AppendToCell(table, table.RowCount - 2, 7, (d.Hours.selfWork + d.Hours.seminar).ToString(), 8);

            AppendToCell(table, table.RowCount - 2, 8, d.Hours.seminar.ToString(), 8);
            AppendToCell(table, table.RowCount - 2, 10, d.Hours.seminar.ToString(), 8);
            AppendToCell(table, table.RowCount - 2, 13, d.Hours.selfWork.ToString(), 8);
            AppendToCell(table, table.RowCount - 2, 14, d.Hours.selfWork.ToString(), 8);

            AppendToCell(table, table.RowCount - 1, (d.IsExam ? 8 : 7), (d.IsExam ? "18" : "4"), 8);
        }
        private static void FillLessonByThemeTable(DocX template, DisciplineInfo d, int discId)
        {
            var table = template.Tables[11 + TablesInDiscTempl * discId];
            for (int i = 0; i < d.Themes.Count; i++)
            {
                table.InsertRow(1);
                AppendToCell(table, 1 + i, 0, i + 1);
            }
            SetTableBorder(table);
        }

        private static void ReplacePlaceholders(DocX template)
        {
            var paragraphs = template.Paragraphs.Where(p => p.Text != "");
            foreach (var p in paragraphs)
                foreach (var repl in TemplateReplacements)
                    p.ReplaceText(repl.Key, repl.Value);
        }

        public static void BuildDocxFromTemplate(ParseInfo pi, string templateFilename)
        {
            if (!Directory.Exists("out"))
                Directory.CreateDirectory("out");
            LoadReplacements(pi);

            using (var d = DocX.Load(templateFilename))
            {
                for (int i = 0; i < pi.Disciplines.Count; i++)
                {
                    d.InsertDocument(DocX.Load("disc_template.docx"));
                    FillEducResults(d, i);
                    FillDiscContentTable(d, pi, i);
                    FillHourTables(d, pi, i);
                }
                ReplacePlaceholders(d);
                d.SaveAs("out\\" + pi.CourseName + ".docx");
            }
        }

        private static void LoadReplacements(ParseInfo pi)
        {

            TemplateReplacements = new Dictionary<string, string>
            {
                { "<MODULE_NAME>", pi.CourseName },
                { "<DISC_NAME>", pi.CourseName }, //no support for multi-discipline modules yet
                { "<YEAR>", DateTime.Now.Year.ToString()},
                { "<M_ZE>", pi.ModuleZe.ToString() },
                { "<CITY>", "Екатеринбург" },
                { "<NAME>", "Иванов И.И." },
                { "<POSITION>", "Искатель интересных историй" },
                { "<UNIVERSITY_NAME>", "ВУЗ им. Иванова И.И." },
                { "<SEMESTER_NO>", "5"},
                { "<TEST_TYPE>", "экзамен"}
            };
            for (int i = 0; i < pi.Disciplines.Count; i++)
                TemplateReplacements.Add("<D" + i + "_ZE>", pi.Disciplines[i].Ze.ToString());
        }
        private static void AppendToCell(Table t, int row, int cell, int value,
            int fontSize = 12, int paragraph = 0)
        {
            AppendToCell(t, row, cell, value.ToString(), fontSize, paragraph);
        }
        private static void AppendToCell(Table t, int row, int cell, string text,
            int fontSize = 12, int paragraph = 0)
        {
            t.Rows[row].Cells[cell].Paragraphs[paragraph].Append(text).FontSize(fontSize);
        }
        private static void SetTableBorder(Table t)
        {
            foreach (TableBorderType borderType in Enum.GetValues(typeof(TableBorderType)))
                t.SetBorder(borderType,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
        }

    }
}