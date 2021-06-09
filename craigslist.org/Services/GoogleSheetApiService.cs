using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace craigslist.org.Services
{
    public static class GoogleSheetApiService
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets, SheetsService.Scope.Drive, SheetsService.Scope.DriveFile };
        static string ApplicationName = "My applications";
        public static SheetsService Service;
        public static UserCredential credential;
        public static async Task<SheetsService> Credential()
        {


            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                const string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine(credPath);
            }

            Service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            })
            {
                HttpClient =
                {
                    Timeout = TimeSpan.FromMinutes(5)
                }
            };
            return Service;

        }

        public static async Task<int> CreateNewSheet(string googleSheetId, string sheetName)
        {
            var addSheetRequest = new AddSheetRequest { Properties = new SheetProperties { Title = sheetName } };
            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            batchUpdateSpreadsheetRequest.Requests.Add(new Request { AddSheet = addSheetRequest });
            var batchUpdateRequest = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, googleSheetId);
            var file = await batchUpdateRequest.ExecuteAsync();
            return file.Replies.First().AddSheet.Properties.SheetId.GetValueOrDefault();
        }

        public static async Task AppendData(List<IList<object>> values, string googleSheetId, string sheetName)
        {
            var request = Service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, googleSheetId, $"{sheetName}!A1:A");
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            await request.ExecuteAsync();

        }

        public static async Task<string> GetRange(string googleSheetId, string sheetName)
        {
            var range = sheetName.Replace(":", "");
            range = range + "!A1:A";
            //try
            //{
            var getRequest = Service.Spreadsheets.Values.Get(googleSheetId, range);

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            var getResponse = await getRequest.ExecuteAsync();
            var getValues = getResponse.Values;
            if (getValues != null)
            {
                var currentCount = getValues.Count() + 1;
                range = sheetName.Replace(":", "") + "!A" + currentCount + ":A";
            }
            else
            {
                range = sheetName.Replace(":", "") + "!A1:A";
            }


            return range;
        }

        public static async Task<int> GetSheetId(string googleSheetId, string sheetName)
        {
            var sheetId = 0;

            var ssRequest = Service.Spreadsheets.Get(googleSheetId);
            var ss = await ssRequest.ExecuteAsync();

            foreach (var sheet in ss.Sheets)
            {
                var title = sheet.Properties.Title;
                if (sheetName.Equals(title))
                {
                    if (sheet.Properties.SheetId != null)
                    {
                        sheetId = (int)sheet.Properties?.SheetId;
                        break;
                    }
                }
            }
            return sheetId;
        }

        public static async Task<bool> CheckIfSheetExist(string googleSheetId, string sheetName)
        {
            var isExist = false;
            //try
            //{
            var ssRequest = Service.Spreadsheets.Get(googleSheetId);
            var ss = await ssRequest.ExecuteAsync();
            foreach (var sheet in ss.Sheets)
            {
                var title = sheet.Properties.Title;
                if (title.Equals(sheetName))
                {
                    isExist = true;
                }
            }
            //}
            //catch (Exception)
            //{

            //    await credential.RefreshTokenAsync(CancellationToken.None);
            //    //Service = new SheetsService(new BaseClientService.Initializer()
            //    //{
            //    //    HttpClientInitializer = credential,
            //    //    ApplicationName = ApplicationName,
            //    //});
            //    await CheckIfSheetExist(googleSheetId, sheetName);
            //}
            return isExist;
        }

        public static async Task DeleteRows(string googleSheetId, string sheetName, int sheetId)
        {
            var request =
                new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Range = new GridRange() { SheetId = sheetId },
                        Fields = "*"
                    }
                };
            var requestContainer = new List<Request> { request };
            var deleteRequest = new BatchUpdateSpreadsheetRequest { Requests = requestContainer };
            var deletion = new SpreadsheetsResource.BatchUpdateRequest(Service, deleteRequest, googleSheetId);
            await deletion.ExecuteAsync();
        }
    }
}
