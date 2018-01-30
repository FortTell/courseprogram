using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using DataClasses;


namespace DocxBuilder
{
    public static class Builder
    {
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
    }
}