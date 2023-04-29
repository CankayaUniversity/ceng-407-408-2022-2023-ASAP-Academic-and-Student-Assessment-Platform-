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
        private const string PathToCredentials = "C:\\Users\\emreh\\Desktop\\ASAP P\\ASAP_Project\\client_secret_714044421228-cugq90i34shjhu5ifs9lmh06fop801ro.apps.googleusercontent.com.json";
        public static async void UploadFile()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.ShowDialog();

            try
            {

            //Solution Explorerda bulunan credentials dosyaları ile adlarını değiştirin
            IDataStore tokenStorage = new FileDataStore("C:\\Users\\emreh\\Desktop\\ASAP P\\ASAP_Project\\SendedAccountCredential.TokenResponse-userName", false);

                

                UserCredential servicecredential;
                await using (var stream = new FileStream(PathToCredentials, FileMode.Open, FileAccess.Read))
                {
                    servicecredential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        new[] {DriveService.ScopeConstants.DriveReadonly},
                        "userName",
                        CancellationToken.None,
                        tokenStorage)
                    .Result;
                }

                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = servicecredential
                });

                var request = service.Files.List();
                var results = await request.ExecuteAsync();

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


            catch (Exception ex)
            {
                MessageBox.Show($"Error uploading file to Google Drive: {ex.Message}");
            }
        }

            //var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            //    new ClientSecrets { ClientId = "606566811129-0v7iesu2r2ehmchfhi56ivf6kuujn7sc.apps.googleusercontent.com", ClientSecret = "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" },
            //    new[] { DriveService.Scope.Drive },
            //    "user",
            //    System.Threading.CancellationToken.None).Result;

            //// Create the Drive service.
            //var service = new DriveService(new BaseClientService.Initializer()
            //{
            //    HttpClientInitializer = credential,
            //    ApplicationName = "ASAP Project"
            //});

            // Upload the selected file to Google Drive.
            //var fileMetadata = new Google.Apis.Drive.v3.Data.File();
            //fileMetadata.Name = System.IO.Path.GetFileName(openFileDialog.FileName);
            //var filePath = openFileDialog.FileName;
            //using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open))
            //{
            //    var uploadRequest = service.Files.Create(fileMetadata, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            //    uploadRequest.Upload();
            //}
        }
    }
