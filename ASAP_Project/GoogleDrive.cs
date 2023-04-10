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

            ServiceAccountCredential credential;

            //A
            // Load the service account credentials from the JSON key file.
            using (var stream = new FileStream("/credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(DriveService.ScopeConstants.Drive)
                    .UnderlyingCredential as ServiceAccountCredential;
            }

            // Create the Drive service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "MyApp",
            });

            // Upload the selected file to Google Drive.
            var fileMetadata = new Google.Apis.Drive.v3.Data.File();
            fileMetadata.Name = System.IO.Path.GetFileName(openFileDialog.FileName);
            var filePath = openFileDialog.FileName;
            using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open))
            {
                var uploadRequest = service.Files.Create(fileMetadata, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                uploadRequest.Upload();
            }
        }
    }
}
