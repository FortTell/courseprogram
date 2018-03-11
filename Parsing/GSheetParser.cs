using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace Parsing
{
    public class GSheetParser
    {
        static string[] scopes = { SheetsService.Scope.Spreadsheets };
        static string appName = "123";
        static string SpreadsheetId = "1lTdVIv9qOw8bLmdnALVs0xc5FBhrtiefnKRPPiGmohc";

        public static void ConnectToSheet()
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
            var response = GetValues("page1!A1:B3", service);
            foreach (var d in response.Values)
                foreach (var c in d)
                    Console.WriteLine(".." + c.ToString() + " ");
            var updValues = new List<object>() { "My Cell Text" };
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "ROWS";
            valueRange.Values = new List<IList<object>> { updValues };

            var update = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, "page1!B1");
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var result2 = update.Execute();

            response = GetValues("page1!A1:B3", service);
            foreach (var d in response.Values)
                foreach (var c in d)
                    Console.WriteLine(".." + c.ToString() + " ");
        }

        public static ValueRange GetValues(string range, SheetsService service)
        {
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            return response;
        }
    }
}
