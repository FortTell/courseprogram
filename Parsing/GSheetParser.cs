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
        static string[] scopes = { SheetsService.Scope.Spreadsheets };
        static string appName = "123";
        static string SpreadsheetId;
        SheetsService Service;

        public GSheetParser(string spreadsheetId = "1lTdVIv9qOw8bLmdnALVs0xc5FBhrtiefnKRPPiGmohc")
        {
            SpreadsheetId = spreadsheetId;
            Service = ConnectToSheetsSvc();
        }

        public void PasteParseInfoToSheet(ParseInfo pi)
        {
            var valuesToUpd = new List<ValueRange>();
            var courseName = new List<IList<object>> { new List<object> { pi.courseName } };
            valuesToUpd.Add(PrepareVR(courseName, "page1!B1"));
            valuesToUpd.Add(PrepareVR(courseName, "page1!B3"));
            valuesToUpd.Add(PrepareVR(new List<IList<object>>
                { pi.themes.Select(t => t.title).ToList<object>()},
                "page1!B7:" + (char)('B' + pi.themes.Count) + "7"));
            valuesToUpd.Add(PrepareVR(new List<IList<object>>
                { pi.themes.Select(t => String.Join(" ",t.topics)).ToList<object>() },
                "page1!B8:" + (char)('B' + pi.themes.Count) + "8"));
            var bur = new SpreadsheetsResource.ValuesResource.BatchUpdateRequest(Service, 
                new BatchUpdateValuesRequest {
                    Data = valuesToUpd,
                    ValueInputOption = "USER_ENTERED"
                }, SpreadsheetId);
            bur.Execute();
        }

        public static SheetsService ConnectToSheetsSvc()
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

        public static ValueRange GetValues(string range, SheetsService service)
        {
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            return response;
        }

        /*            var response = GetValues("page1!A1:B3", service);
            foreach (var d in response.Values)
                foreach (var c in d)
                    Console.WriteLine(".." + c.ToString() + " ");
            UpdateValues(service, new List<IList<object>> { new List<object> { "123" } }, "page1!B2");
        */
    }
}
