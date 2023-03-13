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

            UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "C:\\Users\\emreh\\source\\repos\\ASAP Project\\ASAP Project\\token.json\\";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    clientSecrets: GoogleClientSecrets.Load(stream).Secrets,
                    new[] { DriveService.Scope.Drive },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            var service = new DriveService(new BaseClientService.Initializer()
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
