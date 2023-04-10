using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using static Google.Apis.Drive.v3.DriveService;
using Newtonsoft.Json.Linq;
using static System.Formats.Asn1.AsnWriter;
using System.Net;
using static Google.Apis.Requests.BatchRequest;
using System.IO;
using System.Threading;

using Microsoft.Office.Interop.Excel;
using System.IO;
using Google.Apis.Util;

namespace ASAP_Project
{
    public class GoogleDrive
    {

        public static void UploadFile()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.ShowDialog();




            string clientId = "your_client_id";
            string clientSecret = "your_client_secret";

            string[] scopes = new string[] { DriveService.Scope.Drive };

            UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
             new ClientSecrets
             {
                ClientId = clientId,
                ClientSecret = clientSecret
             },
            scopes,
             "user",
             System.Threading.CancellationToken.None,
             new FileDataStore("Drive.Auth.Store")).Result;

            string accessToken = credential.Token.AccessToken;




            var service = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = new UserCredential(new TokenResponse
                {
                    AccessToken = accessToken
                }),
                ApplicationName = "Your Application Name"
            });



        }
    }
}
