using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace SheetsQuickstart {
    class Program {

        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "TuntisChallenge";
        //MyID and SecretKey
        static String spreadsheetId = "1SKQHuBH1T49my-JT_wNoLcEtZUKSp4vM_GjXvNtz1JI";
        static String sheet = "engenharia_de_software";
        static SheetsService service;

        static void Main(string[] args) {

            Laucher();

            WhiteEntry(ReadEntry());

            //not necessary in this challenge, but I created 
            //UpdateEntry(ReadEntry());
           
        }

        static void Laucher() {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read)) {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        static IList<IList<Object>> ReadEntry() {
            var range = $"{sheet}!C4:F27";

            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();

            IList<IList<Object>> values = response.Values;
            List<IList<Object>> newValue = new List<IList<Object>>();
            IList<Object> value = new List<Object>();

            if (values != null && values.Count > 0) {
                foreach (var row in values) {
                    double media = ((Convert.ToInt32(row[1]) / 10.0) + (Convert.ToInt32(row[2]) / 10.0) + (Convert.ToInt32(row[3]) / 10.0)) / 3;
                    Console.WriteLine($"student average: {media.ToString("F2")}");
                    value = new List<Object>();
                    if (Convert.ToInt32(row[0]) > 15) {
                        value.Add("Reprovado por Falta");
                        value.Add("0");
                    }
                    else if (media >= 5 && media < 7) {
                        value.Add("Exame Final");
                        value.Add(10 - media);
                    }
                    else if (media < 5) {
                        value.Add("Reprovado por Nota");
                        value.Add("0");
                    }
                    else {
                        value.Add("Aprovado");
                        value.Add("0");
                    }
                    newValue.Add(value);
                }
            }
            else {
                Console.WriteLine("No data found.");
            }
            return newValue;
        }

        static void WhiteEntry(IList<IList<Object>> newValue) {
            var range = $"{sheet}!G4:G27";
            var valueRange = new ValueRange();

            valueRange.Values = newValue;

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();
        }

        static void UpdateEntry(IList<IList<Object>> newValue) {
            var range = $"{sheet}!G4:G27";
            var valueRange = new ValueRange();

            valueRange.Values = newValue;

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();

        }

    }
}
