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

            // EMRE FUCKING DID THIS//
            //

            string accesstoken = "ya29.a0AVvZVsqGpBbpYwDdAMLkX18R8Ntqcq-rYKhry1f1PnRr8_uNRlrjmIOJ6u0dHehoqKd3PihenLZUVlNehMJuwEHQWeiXacwYgm8NF3cudAZJ2kF5oCusa_lzxsiYuohBvabiy_bkWXiylgoNyAPlYLTPi9JbNHG4aCgYKAV0SARASFQGbdwaIrFk78VrafIOMklUo5C4EbA0167";


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
