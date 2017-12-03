using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using System.IO;
using System.Drawing;

namespace CourseProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var d = DocX.Create("123.docx"))
            {
                var t = d.InsertTable(3, 3);
                t.SetBorder(TableBorderType.Top, new Border(BorderStyle.Tcbs_double, BorderSize.one, 1, Color.Black));
                t.Rows[0].Cells[1].Paragraphs[0].Append("123");
                d.Save();
            }
            Console.WriteLine("!");
        }
    }
}
