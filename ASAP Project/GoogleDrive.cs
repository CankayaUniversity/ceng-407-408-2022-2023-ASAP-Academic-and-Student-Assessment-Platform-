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

namespace ASAP_Project
{
    public class GoogleDrive
    {
        public static void UploadFile()
        {
            string oSelectedFile = "";
            System.Windows.Forms.OpenFileDialog oDlg = new System.Windows.Forms.OpenFileDialog();
            if (System.Windows.Forms.DialogResult.OK == oDlg.ShowDialog())
            {
                oSelectedFile = oDlg.FileName;

            }

            string clientId = "606566811129-vqo2k7m0nn2h9g80982a2p61gv9fsrb9.apps.googleusercontent.com";
            string clientSecret = "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";                       

            var scopes = new[] { DriveService.Scope.Drive };

            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets { ClientId = clientId, ClientSecret = clientSecret },
                scopes,
                Environment.UserName,
                CancellationToken.None,
                new FileDataStore("Drive.Auth.Store")).Result;

            if (credential == null)
            {
                MessageBox.Show("Unable to authorize access to Google Drive API.");
                return;
            }

            /*UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {

                string credPath = "..\\ASAP Project\\ASAP Project\\acc_token\\";

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    clientSecrets: GoogleClientSecrets.Load(stream).Secrets,
                    new[] { DriveService.Scope.Drive },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }*/

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
            using (var stream = new FileStream(oSelectedFile, FileMode.Open))
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
