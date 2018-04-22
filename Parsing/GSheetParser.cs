using DataClasses;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using VR = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace Parsing
{
    public class GSheetParser
    {
        GSheetApiHandler Handler;

        public GSheetParser(GSheetApiHandler handler)
        {
            Handler = handler;
        }

        public List<ValueRange> PrepareInfoForPasting(ParseInfo pi, int discId)
        {
            var valuesToUpd = new List<ValueRange>();
            var disc = pi.Disciplines[discId];
            var courseName = new List<IList<object>> { new List<object> { pi.CourseName } };
            valuesToUpd.Add(PrepareVR(courseName, "page1!B1"));
            valuesToUpd.Add(PrepareVR(courseName, "page1!B3"));
            valuesToUpd.Add(PrepareVR(new List<IList<object>>
                { disc.Themes.Select(t => t.title).ToList<object>()},
                "page1!B7:" + (char)('B' + disc.Themes.Count) + "7"));
            valuesToUpd.Add(PrepareVR(new List<IList<object>>
                { disc.Themes.Select(t => String.Join(". ", t.topics)).ToList<object>() },
                "page1!B8:" + (char)('B' + disc.Themes.Count) + "8"));
            return valuesToUpd;
        }

        public ParseInfo ParseInfoFromSheet()
        {
            var dInfo = Handler.GetValues("page1!A5:F5");
            var dName = Handler.GetValues("page1!B3");
            var dThemes = Handler.GetValues("page1!B7:F8");
            var pi = new ParseInfo
            {
                CourseName = Handler.GetValues("page1!B1").Values[0][0].ToString(),
                Disciplines = new List<DisciplineInfo> {
                    DisciplineInfo.CreateSecondPassDI(
                        dName.Values[0][0].ToString(),
                        int.Parse(dInfo.Values[0][0].ToString()),
                        (int.Parse(dInfo.Values[0][3].ToString()),
                            int.Parse(dInfo.Values[0][4].ToString()),
                            int.Parse(dInfo.Values[0][5].ToString())),
                        GetThemeInfos(dThemes),
                        dInfo.Values[0][1].ToString().Contains("Экзамен")
                    )
                }
            };
            return pi;
        }

        private List<(string, List<string>)> GetThemeInfos(ValueRange vr)
        {
            return vr.Values[0].Zip(vr.Values[1],
                (ti, to) => (ti.ToString(), to.ToString().Split(". ").ToList())).ToList();
        }

        public ValueRange PrepareVR(List<IList<object>> updatedValues, string range, bool pasteInRows = true)
        {
            return new ValueRange
            {
                MajorDimension = pasteInRows ? "ROWS" : "COLUMNS",
                Values = updatedValues,
                Range = range
            };
        }
    }
}
