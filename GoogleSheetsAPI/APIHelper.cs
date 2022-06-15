using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GoogleSheetsAPI
{
    public class APIHelper
    {
        public SheetsService service { get; set; }


        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string APPLICATION_NAME = "Google Sheets API";


        public APIHelper()
        {
            InitializeService();
        }

        private void InitializeService()
        {
            UserCredential credential = GetCredentialsFromFile();

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = APPLICATION_NAME
            });
        }

        private UserCredential GetCredentialsFromFile()
        {
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                       GoogleClientSecrets.FromStream(stream).Secrets,
                       Scopes,
                       "user",
                       CancellationToken.None,
                       new FileDataStore(credPath, true)).Result;

                return credential;
            }
        }

    }
}
