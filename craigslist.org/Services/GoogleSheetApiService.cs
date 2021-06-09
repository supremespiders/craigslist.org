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

        public static async Task<string> CreateNewSheet(string googleSheetId, string sheetName)
        {
            var file = new BatchUpdateSpreadsheetResponse();
            //try
            //{
            var addSheetRequest = new AddSheetRequest { Properties = new SheetProperties { Title = sheetName.Replace(":", "") } };
            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            batchUpdateSpreadsheetRequest.Requests.Add(new Request { AddSheet = addSheetRequest });
            var batchUpdateRequest = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, googleSheetId);
            file = await batchUpdateRequest.ExecuteAsync();
            return file.SpreadsheetId;
            //}
            //catch (Exception)
            //{

            //    await credential.RefreshTokenAsync(CancellationToken.None);
            //    //Service = new SheetsService(new BaseClientService.Initializer()
            //    //{
            //    //    HttpClientInitializer = credential,
            //    //    ApplicationName = ApplicationName,
            //    //});
            //    await CreateNewSheet(googleSheetId, sheetName);
            //}
            //return file.SpreadsheetId;
        }

        public static async Task AppendData(List<IList<object>> values, string googleSheetId, string sheetName)
        {

            //var range = await GetRange(googleSheetId, sheetName);
            var range = sheetName.Replace(":", "");
            range = range + "!A1:A";
            var request =
                Service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, googleSheetId, range);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            //request.ResponseDateTimeRenderOption
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

        public static async Task DeleteRows(string googleSheetId, string sheetName, string sheetId)
        {
            var range = await GetRange(googleSheetId, sheetName);
            var endIndex = 0;
            var endIndexBuilder = new StringBuilder();
            var x = range.IndexOf('!') + 2;
            for (var i = x; i < range.Length; i++)
            {
                var ch = range[i];
                if (char.IsDigit(ch))
                {
                    endIndexBuilder.Append(ch);
                    continue;
                }
                break;
            }

            endIndex = int.Parse(endIndexBuilder.ToString());
            var request =
                new Request
                {
                    DeleteDimension =
                        new DeleteDimensionRequest
                        {
                            Range =
                                new DimensionRange {  SheetId = sheetId, Dimension = "ROWS", StartIndex = 0, EndIndex = endIndex }
                        }/*, DeleteRange = new DeleteRangeRequest { Range = new GridRange { } }*/
                };
            var requestContainer = new List<Request> { request };
            var deleteRequest = new BatchUpdateSpreadsheetRequest { Requests = requestContainer };
            var deletion = new SpreadsheetsResource.BatchUpdateRequest(Service, deleteRequest, googleSheetId);
            await deletion.ExecuteAsync();

        }


    }
}
