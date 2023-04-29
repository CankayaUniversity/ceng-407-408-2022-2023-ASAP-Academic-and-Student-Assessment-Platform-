using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;

namespace ASAP_Project
{
    public class GoogleDrive
    {

        public static void UploadFile()
        {
            //Bu PCde son açılan google accountuna yüklemekte kodu,
            //Biz bunun asap driveına yüklenmesini sağlamalıyız
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.ShowDialog();

            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets { ClientId = "606566811129-0v7iesu2r2ehmchfhi56ivf6kuujn7sc.apps.googleusercontent.com", ClientSecret = "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" },
                new[] { DriveService.Scope.Drive },
                "asaproject2023@gmail.com",
                CancellationToken.None,
                new FileDataStore("Drive.Auth.Store")).Result;

            // Create the Drive service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "ASAP Project"
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
