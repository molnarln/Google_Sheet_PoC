using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Google_Sheet_TestProject.Controllers
{
    public class GoogleSheetController : Controller
    {
        [HttpGet("googlesheet")]
        public  IActionResult GetGoogleSheetAsync()
        {

            //Profile and email scopes are arbitrary, just for the sake of the test
            //If delete token.json, sometimes need to open the endpoint twice, or clean and build solution.
            //But it works!!!
            string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly, "profile", "email" };
            string ApplicationName = "Google Sheets API .NET Quickstart";

            UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential =  GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "laszlo.molnar25@gmail.com",
                    CancellationToken.None,
                    new FileDataStore("token.json", true)).GetAwaiter().GetResult();
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            string spreadsheetId = "1h4M6fKsF1hOFYWVfuQ4jPFR_ZhYnuLZpBF2C9Uk1GlQ";
            string range = "Students!A2:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }

            return Ok(values);

        }
    }
}
