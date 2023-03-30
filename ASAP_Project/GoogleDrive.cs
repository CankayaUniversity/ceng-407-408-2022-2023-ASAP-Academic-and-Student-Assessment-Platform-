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
using Dropbox.Api;
using Microsoft.Office.Interop.Excel;
using System.IO;
using Dropbox.Api.Files;

namespace ASAP_Project
{
    public class GoogleDrive
    {

        public static void UploadFile()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.ShowDialog();
            // EMRE FUCKING DID THIS//
            //

            string accesstoken = "accesstoken";
            var credential = GoogleCredential.FromAccessToken(accesstoken);


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
