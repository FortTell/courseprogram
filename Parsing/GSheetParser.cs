﻿using DataClasses;
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
        static string[] scopes = { SheetsService.Scope.Spreadsheets };
        static string appName = "123";
        static string SpreadsheetId;
        SheetsService Service;

        public GSheetParser(string spreadsheetId = "1lTdVIv9qOw8bLmdnALVs0xc5FBhrtiefnKRPPiGmohc")
        {
            SpreadsheetId = spreadsheetId;
            Service = ConnectToSheetsSvc();
        }

        public void PasteParseInfoToSheet(ParseInfo pi, int discId)
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
            var bur = new SpreadsheetsResource.ValuesResource.BatchUpdateRequest(Service,
                new BatchUpdateValuesRequest
                {
                    Data = valuesToUpd,
                    ValueInputOption = "USER_ENTERED"
                }, SpreadsheetId);
            bur.Execute();
        }
        public ParseInfo ParseInfoFromSheet()
        {
            var di = GetValues("page1!B3:B6");
            var dThemes = GetValues("page1!B7:F8");
            var pi = new ParseInfo
            {
                CourseName = GetValues("page1!B1").Values[0][0].ToString(),
                Disciplines = new List<DisciplineInfo> {
                    new DisciplineInfo
                    {
                        Name = di.Values[0][0].ToString(),
                        Ze = int.Parse(di.Values[1][0].ToString()),
                        PracticeHours = int.Parse(di.Values[2][0].ToString()),
                        IsExam = di.Values[3][0].ToString() == "Экзамен",
                        Themes = GetThemeInfos(dThemes)
                    }
                }
            };
            return pi;
        }

        public SheetsService ConnectToSheetsSvc()
        {
            GoogleCredential credential;
            string credPath = Path.Combine(new DirectoryInfo(System.Environment.CurrentDirectory).Parent.FullName, "Parsing");
            using (var stream = new FileStream(Path.Combine(credPath, "client_secret.json"),
                FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
                Console.WriteLine("Credential file saved to: " + Path.Combine(credPath, "cred.json"));
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = appName
            });
            return service;
        }

        public void HAX(ParseInfo pi)
        {
            var uspr = new UpdateDimensionPropertiesRequest();
            uspr.Properties.PixelSize = pi.Disciplines[0].Themes[0].topics.Sum(t => t.Length + 2);
            uspr.Range = new DimensionRange { Dimension = "ROWS", SheetId = 0, StartIndex = 1, EndIndex = 6 + 1 };
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
        public UpdateValuesResponse UpdateValues(List<IList<object>> updatedValues, string range,
            bool pasteAsRaw = false, bool pasteInRows = true)
        {
            var vr = PrepareVR(updatedValues, range, pasteInRows);
            var updateReq = Service.Spreadsheets.Values.Update(vr, SpreadsheetId, range);
            updateReq.ValueInputOption = pasteAsRaw ?
                VR.UpdateRequest.ValueInputOptionEnum.RAW :
                VR.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            return updateReq.Execute();
        }

        public ValueRange GetValues(string range)
        {
            var request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            return response;
        }
    }
}
