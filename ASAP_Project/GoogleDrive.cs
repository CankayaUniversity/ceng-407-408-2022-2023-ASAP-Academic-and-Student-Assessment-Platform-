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
            //Sonrası silinebilir ve baştan yazılmaya açık, ama bi ara çalışıyordu
            // EMRE FUCKING DID THIS//
            //

            //string clientId = "606566811129-0v7iesu2r2ehmchfhi56ivf6kuujn7sc.apps.googleusercontent.com";
            //string clientSecret = "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";
            ////string[] scope = { "https://www.googleapis.com/auth/drive.file" };
            //string refreshToken = "1//04dECzas1BhGNCgYIARAAGAQSNwF-L9IrqkbzmoLGSyjrH03u6YIfjraGviDkd0Kj4Tr13tViHgCQeC87IXtXEIr5TwQ7C0CGQow";

            //UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            //    new ClientSecrets
            //    {
            //        ClientId = clientId,
            //        ClientSecret = clientSecret
            //    },
            //    new[] { DriveService.Scope.Drive },
            //    "user",
            //     System.Threading.CancellationToken.None,
            //     new Google.Apis.Util.Store.FileDataStore("Drive.Api.Auth.Store")).Result;

            //credential.Token = new TokenResponse
            //{
            //    RefreshToken = refreshToken
            //};

            //bool success = credential.RefreshTokenAsync(CancellationToken.None).Result;
            //string accessToken = credential.Token.AccessToken;

            DriveService service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "MyConsoleApp",
            });

            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = "TEST"
            };




            FilesResource.CreateMediaUpload request;
            using (var stream = new FileStream(openFileDialog.FileName, FileMode.Open))
            {
                request = service.Files.Create(
                    fileMetadata, stream, "application/vnd.ms-excel");
                request.Fields = "id";
                request.Upload();
            }
            var file = request.ResponseBody;
        }
    }
}
