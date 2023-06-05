using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Windows;

namespace ASAP_Project
{
    public class GoogleDrive
    {
        //Solution Explorerda bulunan credentials dosyaları ile adlarını değiştirin    
        static IDataStore tokenStorage = new FileDataStore(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SendedAccountCredential"), false);

        public static async void UploadCourse(string filepath)
        {
            try
            {

                string folderid = "1yaDOAB2U008ohDirn03H8-RB1r6LJoFc";
                string filePath = filepath;

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
                var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = Path.GetFileName(filePath),
                    Parents = new[] { folderid }
                };
                FilesResource.CreateMediaUpload request;
                using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open))
                {
                    request = service.Files.Create(fileMetadata, stream, "application/vnd.ms-excel");
                    request.Fields = "id";
                    request.Upload();
                }
                var file = request.ResponseBody;
            }


            catch (Exception ex)
            {
                MessageBox.Show($"Error uploading file to Google Drive: {ex.Message}");
            }
        }


        public static async void UploadFile()
        {
            
            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.ShowDialog();

                           

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

        public static System.IO.MemoryStream GetFile(string name)
        {
            try
            {

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

                var request = service.Files.List();
                request.Q = "name='" + name + "' and trashed = false";
                request.Fields = "nextPageToken, files(id)";
                var results = request.Execute().Files;

                if (results == null || results.Count == 0)
                {
                    MessageBox.Show("No files found.");
                }

                var file = service.Files.Get(results[0].Id).Execute();

                var downloadfile = service.Files.Get(results[0].Id);
                var stream = new MemoryStream();
                downloadfile.Download(stream);

                /*OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Folders|*.none";
                openFileDialog.CheckFileExists = false;
                openFileDialog.CheckPathExists = true;
                openFileDialog.FileName = name;
                string filePath = null;
                if (openFileDialog.ShowDialog() == true)
                {
                    filePath = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
                }
                
                

                using (var fileStream = new FileStream(openFileDialog.FileName, FileMode.Create, FileAccess.Write))
                {
                    stream.WriteTo(fileStream);
                }*/

                return stream;
            }


            catch (Exception ex)
            {
                MessageBox.Show($"Error downloading file to Google Drive: {ex.Message}");
                return null;
            }
        }

        public static List<string> course_list = new List<string>();

        public static List<string> gen_excel = new List<string>();
        public static List<string> getCourseList()
        {
            course_list.Clear();

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
            string folderId = "1yaDOAB2U008ohDirn03H8-RB1r6LJoFc";
            var fileListRequest = service.Files.List();
            fileListRequest.Q = $"'{folderId}' in parents";
            fileListRequest.PageSize = 10; // Set the number of files to retrieve per page
            fileListRequest.Fields = "nextPageToken, files(name, id, mimeType)"; // Specify the fields to retrieve
            var fileList = fileListRequest.Execute();

            foreach (var file in fileList.Files)
            {
                course_list.Add(file.Name);
            }

            return course_list;
        }

        public static List<string> getGenExcelList()
        {
            gen_excel.Clear();

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
            string folderId = "1CISiRSjdUMB8h3v64CNZgc9bCeVfuIKH";
            var fileListRequest = service.Files.List();
            fileListRequest.Q = $"'{folderId}' in parents";
            fileListRequest.PageSize = 10; // Set the number of files to retrieve per page
            fileListRequest.Fields = "nextPageToken, files(name, id, mimeType)"; // Specify the fields to retrieve
            var fileList = fileListRequest.Execute();

            foreach (var file in fileList.Files)
            {
                gen_excel.Add(file.Name);
            }

            return gen_excel;
        }

        public static void DeleteFile(string filepath)
        {
            string folderId = "1yaDOAB2U008ohDirn03H8-RB1r6LJoFc";
            string filePath = filepath;

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

            var request = service.Files.List();
            request.Q = $"name = '{filepath}'";
            var result = request.Execute();

            service.Files.Delete(result.Files[0].Id).Execute();


        }

    }
}
