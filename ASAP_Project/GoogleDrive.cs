using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;

namespace ASAP_Project
{
    public class GoogleDrive
    {
        //Solution Explorerda bulunan credentials dosyaları ile adlarını değiştirin
        public static async void UploadFile()
        {
            

            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.ShowDialog();

                //Solution Explorerda bulunan credentials dosyaları ile adlarını değiştirin
                var tokenStorage = new FileDataStore("C:\\Users\\hayre\\Source\\Repos\\ceng-407-408-2022-2023-ASAP-Academic-and-Student-Assessment-Platform-\\ASAP_Project\\SendedAccountCredential\\", false);

                var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    new ClientSecrets { ClientId = "714044421228-cugq90i34shjhu5ifs9lmh06fop801ro.apps.googleusercontent.com", ClientSecret = "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" },
                    new[] { DriveService.Scope.Drive },
                    "user",
                    System.Threading.CancellationToken.None,
                    tokenStorage).Result;

                // Create the Drive service.
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "ASAP Project"
                });

                //Upload the selected file to Google Drive.
                var fileMetadata = new Google.Apis.Drive.v3.Data.File();
                fileMetadata.Name = System.IO.Path.GetFileName(openFileDialog.FileName);
                var filePath = openFileDialog.FileName;
                using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open))
                {
                    var uploadRequest = service.Files.Create(fileMetadata, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    uploadRequest.Upload();
                }
               
            }


            catch (Exception ex)
            {
                MessageBox.Show($"Error uploading file to Google Drive: {ex.Message}");
            }
        }

        public static async void GetFile()
        {


            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.ShowDialog();

                //Solution Explorerda bulunan credentials dosyaları ile adlarını değiştirin
                var tokenStorage = new FileDataStore("C:\\Users\\hayre\\Source\\Repos\\ceng-407-408-2022-2023-ASAP-Academic-and-Student-Assessment-Platform-\\ASAP_Project\\SendedAccountCredential\\", false);

                var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    new ClientSecrets { ClientId = "714044421228-cugq90i34shjhu5ifs9lmh06fop801ro.apps.googleusercontent.com", ClientSecret = "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" },
                    new[] { DriveService.Scope.Drive },
                    "user",
                    System.Threading.CancellationToken.None,
                    tokenStorage).Result;

                // Create the Drive service.
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "ASAP Project"
                });

                //Upload the selected file to Google Drive.
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.PageSize = 10;
                listRequest.Fields = "nextPageToken, files(id, name)";

                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
                MessageBox.Show("Files:");
                if (files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        MessageBox.Show("{0} ({1})", file.Name);
                    }
                }
                else
                {
                    MessageBox.Show("No files found.");
                }
            }


            catch (Exception ex)
            {
                MessageBox.Show($"Error uploading file to Google Drive: {ex.Message}");
            }
        }
    }
}
