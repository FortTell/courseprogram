using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using VR = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DataClasses;

namespace Parsing
{
    public class GSheetApiHandler
    {
        public GSheetParser Parser;
        static string[] scopes = { SheetsService.Scope.Spreadsheets };
        static string appName = "123";
        public string SpreadsheetId;
        public string SheetLink { get => "https://docs.google.com/spreadsheets/d/" + SpreadsheetId; }
        public SheetsService Service;

        public GSheetApiHandler(string spreadsheetId = "1lTdVIv9qOw8bLmdnALVs0xc5FBhrtiefnKRPPiGmohc")
        {
            SpreadsheetId = spreadsheetId;
            Service = ConnectToSheetsSvc();
        }

        public void InitParser()
        {
            Parser = new GSheetParser(this);
        }

        public SheetsService ConnectToSheetsSvc()
        {
            GoogleCredential credential;
            string credPath = Path.Combine(new DirectoryInfo(System.Environment.CurrentDirectory).Parent.FullName, "Parsing");
            using (var stream = new FileStream(Path.Combine(credPath, "client_secret.json"),
                FileMode.Open, FileAccess.Read))
                credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = appName
            });
            return service;
        }

        public void PasteInfoToSheet(ParseInfo pi, int discId)
        {
            var bur = new VR.BatchUpdateRequest(
                Service,
                new BatchUpdateValuesRequest
                {
                    Data = Parser.PrepareInfoToPasteToSheet(pi, discId),
                    ValueInputOption = "USER_ENTERED"
                },
                SpreadsheetId);
            bur.Execute();
        }

        public void CopyDiscTemplate(int i)
        {
            var src = new GridRange
            {
                SheetId = 0,
                StartColumnIndex = 0,
                StartRowIndex = 2,
                EndColumnIndex = 5 + 1,
                EndRowIndex = 7 + 1
            };
            var dst = new GridRange
            {
                SheetId = 0,
                StartColumnIndex = 0,
                StartRowIndex = 2 + 8 * i,
                EndColumnIndex = 5 + 1,
                EndRowIndex = 7 + 2 * i + 1
            };
            CopyPasteRequest cpr = new CopyPasteRequest { Source = src, Destination = dst, PasteType = "PASTE_NORMAL" };
            var busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { new Request { CopyPaste = cpr } },
            };
            var r = new SpreadsheetsResource.BatchUpdateRequest(Service, busr, SpreadsheetId);
            r.Execute();
        }

        public UpdateValuesResponse UpdateValues(List<IList<object>> updatedValues, string range,
            bool pasteAsRaw = false, bool pasteInRows = true)
        {
            var vr = Parser.PrepareVR(updatedValues, range, pasteInRows);
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
