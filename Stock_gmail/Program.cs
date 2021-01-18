using System;
using System.IO;
using System.Data;
using System.Configuration;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using ExcelDataReader;
using System.Net;
using System.Text;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using System.Threading.Tasks;

namespace Stock_gmail
{
    class Program
    {

        //In order to use this script we have to create a credentials.json file to connect us to the Google API.
        //Without this we cannot download the google drive files.

        static long files = 0;
        static long directories = 0;

        static string copypath = ConfigurationManager.AppSettings.Get("CopyPath");
        static string localpastepath = ConfigurationManager.AppSettings.Get("LocalPastePath");
        static string serverpastepath = ConfigurationManager.AppSettings.Get("ServerPastePath");
        static string individualdrectory = ConfigurationManager.AppSettings.Get("IndividualDirectory");

        static string strUser = ConfigurationManager.AppSettings.Get("ftpUser");
        static string strPassword = ConfigurationManager.AppSettings.Get("ftpPassword");
        static string strServer = ConfigurationManager.AppSettings.Get("ftpServer");

        static string[] Scopes = { DriveService.Scope.DriveReadonly };
        static string ApplicationName = "Drive API .NET Quickstart";


        static void Main(string[] args)
        {

            GetFilesFromDrive();

            DirectoryInfo di = new DirectoryInfo(copypath);

            try
            {
                // Create a new DirectoryInfo object.
                DirectoryInfo dir = new DirectoryInfo(copypath);

                // Call the GetFileSystemInfos method.
                FileSystemInfo[] infos = dir.GetFileSystemInfos();

                // Pass the result to the ListDirectoriesAndFiles
                // method defined below.
                ListDirectoriesAndFiles(infos);
            }
            catch (Exception e)
            {
                //Console.WriteLine(e.Message);
            }
            finally
            {
                //Console.ReadLine();
            }

            string file_name = "";
            string filepath = "";
            string finallocalpastepath = "";
            string finalserverpastepath = "";
            string fileserverpath = "";
            string finalmovingpath = "";

            foreach (var fi in di.GetFiles("*.xls*"))
            {
                file_name = fi.Name.ToString();
                filepath = copypath + file_name;
                string finalfilename = Path.ChangeExtension(file_name, ".csv");
                finallocalpastepath = localpastepath + finalfilename;
                finalserverpastepath = serverpastepath + finalfilename;

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (FileStream stream1 = System.IO.File.Open(filepath, FileMode.Open, FileAccess.Read))

                {
                    IExcelDataReader reader1;

                    if (file_name.Substring(file_name.Length - 4).Equals(".xls"))
                    {
                        reader1 = ExcelReaderFactory.CreateBinaryReader(stream1);
                    }
                    else
                    {
                        reader1 = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream1);
                    }

                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true,

                            FilterColumn = (columnReader, columnIndex) =>
                            {
                                string header = columnReader.GetString(columnIndex);
                                return (header == "Código Barras" ||
                                        header == "Existencias" ||
                                        header == "En stock" ||
                                        header == "Codigo" ||
                                        header == "CODI"
                                        );
                            },
                        }
                    };

                    var dataSet = reader1.AsDataSet(conf);
                    var dataTable = dataSet.Tables[0];

                    try
                    {
                        // Create the CSV file to which grid data will be exported.
                        FileStream fs = new FileStream(finallocalpastepath, FileMode.Create, FileAccess.ReadWrite);
                        StreamWriter sw = new StreamWriter(fs);
                        // First we will write the headers.
                        int iColCount = dataTable.Columns.Count;
                        for (int i = 0; i < iColCount; i++)
                        {
                            sw.Write(dataTable.Columns[i]);
                            if (i < iColCount - 1)
                            {
                                sw.Write(";");
                            }
                        }
                        sw.Write(sw.NewLine);

                        //Now write all the rows.
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            if (dr[0].ToString() != "")
                            {
                                for (int i = 0; i < iColCount; i++)
                                {
                                    // Change Text "Consultar" to "-1" and Text "0" to "0".
                                    var print = dr[i].ToString();
                                    if (dr[i].ToString() == "Consultar")
                                    {
                                        if (!Convert.IsDBNull(dr[i]))
                                        {
                                            print = "-1";
                                            sw.Write(print);
                                        }
                                        if (i < iColCount - 1)
                                        {
                                            sw.Write(";");
                                        }
                                    }
                                    else
                                    {
                                        if (!Convert.IsDBNull(dr[i]))
                                        {
                                            char MyChar = '+';
                                            string NewString = print.Trim(MyChar);
                                            sw.Write(NewString);
                                        }
                                        if (i < iColCount - 1)
                                        {
                                            sw.Write(";");
                                        }
                                    }
                                }
                            }
                            if (dr[0].ToString() != "")
                            {
                                sw.Write(sw.NewLine);
                            }
                        }
                        sw.Close();
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

                //Here goes the Folder name (in this case, the company name)
                string name1 = "DEUTER";
                string name2 = "ORTOVOX";
                string name3 = "SPORTLAST";
                string name4 = "Others";

                //--- HERE --- Add new name
                //string name4 = "";

                string fileremotepath = "";
                string folderDirectory = "";

                //----HERE ---- 
                //if new name added, copy following code and change variable nameX to the new one created. 

                /***/
                //else if (finallocalpastepath.Contains(nameX) == true)
                //{
                //    finalserverpastepath = serverpastepath + "nameX" + "\\";
                //    DirectoryInfo directory = Directory.CreateDirectory(individualdrectory + nameX + "\\Sincro");
                //    folderDirectory = directory.ToString();
                //    fileremotepath = directory.ToString() + "\\" + finalfilename;
                //    fileserverpath = finalserverpastepath + file_name;

                //}
                /***/

                if (finallocalpastepath.Contains(name1) == true)
                {
                    finalserverpastepath = serverpastepath + "" + "\\";
                    DirectoryInfo directory = Directory.CreateDirectory(individualdrectory+ ""+ "\\Sincro");
                    folderDirectory = directory.ToString();
                    fileremotepath = directory.ToString() + "\\" + finalfilename;
                    fileserverpath = finalserverpastepath + finalfilename;
                    System.IO.File.Delete(filepath);

                }
                else if (finallocalpastepath.Contains(name2) == true)
                {
                    finalserverpastepath = serverpastepath + "" + "\\";
                    DirectoryInfo directory = Directory.CreateDirectory(individualdrectory + "" + "\\Sincro");
                    folderDirectory = directory.ToString();
                    fileremotepath = directory.ToString() + "\\" + finalfilename;
                    fileserverpath = finalserverpastepath + finalfilename;
                    System.IO.File.Delete(filepath);

                }
                else if (finallocalpastepath.Contains(name4) == true)
                {
                    finalserverpastepath = serverpastepath + name4 + "\\";
                    DirectoryInfo directory = Directory.CreateDirectory(individualdrectory + name4 + "\\Sincro");
                    folderDirectory = directory.ToString();
                    fileremotepath = directory.ToString() + "\\" + finalfilename;
                    fileserverpath = finalserverpastepath + finalfilename;
                    System.IO.File.Delete(filepath);
                }
                else
                {
                    finalserverpastepath = serverpastepath + "" + "\\Sincro";
                    DirectoryInfo directory = Directory.CreateDirectory(individualdrectory + "" + "\\Sincro");
                    folderDirectory = directory.ToString();
                    fileremotepath = directory.ToString() + "\\" + finalfilename;
                    fileserverpath = finalserverpastepath + finalfilename;
                    System.IO.File.Delete(filepath);
                }

                if (System.IO.File.Exists(fileremotepath))
                {
                    // This path is a file
                    //ProcessFile(path);
                    //Console.WriteLine("File exists");
                }
                else if (Directory.Exists(folderDirectory))
                {
                    System.IO.File.Move(finallocalpastepath, fileremotepath);
                    System.IO.File.Delete(finallocalpastepath);
                    //Console.WriteLine("File moved");
                }
                else
                {
                    //Console.WriteLine("{0} is not a valid file or directory.", folderDirectory);
                }

                //////Create directory 
                //try
                //{
                //    WebRequest request = WebRequest.Create(string.Format("ftp://{0}/{1}", strServer, finalserverpastepath));
                //    request.Method = WebRequestMethods.Ftp.MakeDirectory;
                //    request.Credentials = new NetworkCredential(strUser, strPassword);
                //    using (var resp = (FtpWebResponse)request.GetResponse())
                //    {
                //        //Console.WriteLine(resp.StatusCode);
                //    }
                //}
                //catch (WebException ex)
                //{
                //    //FtpWebResponse response = (FtpWebResponse)ex.Response;
                //    //if (response.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                //    //{

                //    //}
                //}

                ////hacer request y postear
                //FtpWebRequest ftpRequest;

                //// Crea el objeto de conexión del servidor FTP
                //ftpRequest = (FtpWebRequest)WebRequest.Create(string.Format("ftp://{0}/{1}", strServer, fileserverpath));
                //// Asigna las credenciales
                //ftpRequest.Credentials = new NetworkCredential(strUser, strPassword);
                //// Asigna las propiedades
                //ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
                //ftpRequest.UsePassive = true;
                //ftpRequest.UseBinary = true;
                //ftpRequest.KeepAlive = false;

                //StreamReader sourceStream = new StreamReader(fileremotepath);
                //byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                //sourceStream.Close();
                //ftpRequest.ContentLength = fileContents.Length;

                //Stream requestStream = ftpRequest.GetRequestStream();
                //requestStream.Write(fileContents, 0, fileContents.Length);
                //requestStream.Close();

                //FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();

                //string[] folderfiles = Directory.GetFiles(copypath);
                //foreach (string file in folderfiles)
                //{
                //    System.IO.File.Delete(file);
                //    //Console.WriteLine($"{file} is deleted.");
                //}
            }
        }

            static void ListDirectoriesAndFiles(FileSystemInfo[] FSInfo)
            {
                // Check the FSInfo parameter.
                if (FSInfo == null)
                {
                    throw new ArgumentNullException("FSInfo");
                }

                // Iterate through each item.
                foreach (FileSystemInfo i in FSInfo)
                {
                    // Check to see if this is a DirectoryInfo object.
                    if (i is DirectoryInfo)
                    {
                        // Add one to the directory count.
                        directories++;

                        // Cast the object to a DirectoryInfo object.
                        DirectoryInfo dInfo = (DirectoryInfo)i;

                        // Iterate through all sub-directories.
                        ListDirectoriesAndFiles(dInfo.GetFileSystemInfos());
                    }
                    // Check to see if this is a FileInfo object.
                    else if (i is FileInfo)
                    {
                        // Add one to the file count.
                        files++;
                    }
                }
            }

            static void GetFilesFromDrive()
            {
                UserCredential credential;

                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

                // Create Drive API service.
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Define parameters of request.
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.PageSize = 400;
                listRequest.Fields = "nextPageToken, files(id, name)";

                // List files.
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                    .Files;
                //Console.WriteLine("Files:");

                String path = copypath;

                if (files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        var jetStream = new System.IO.MemoryStream();

                        //Console.WriteLine("{0} ({1})", file.Name, file.Id);
                        FilesResource.GetRequest request = new FilesResource.GetRequest(service, file.Id);
                        //ExportRequest(service, file.Id, "application/vnd.google-apps.file");
                        //(service, file.Id);
                        string pathforfiles = path + file.Name.ToString();

                        request.MediaDownloader.ProgressChanged += (Google.Apis.Download.IDownloadProgress progress) =>
                        {
                            switch (progress.Status)
                            {
                                case Google.Apis.Download.DownloadStatus.Downloading:
                                    {
                                    //Console.WriteLine(progress.BytesDownloaded);
                                    break;
                                    }
                                case Google.Apis.Download.DownloadStatus.Completed:
                                    {
                                        //Console.WriteLine("Download complete.");
                                        using (System.IO.FileStream file = new System.IO.FileStream(pathforfiles, System.IO.FileMode.Create, System.IO.FileAccess.Write))
                                        {
                                            jetStream.WriteTo(file);
                                        }
                                        break;
                                    }
                                case Google.Apis.Download.DownloadStatus.Failed:
                                    {
                                        //Console.WriteLine("Download failed.");
                                        break;
                                    }
                            }
                        };

                        request.DownloadWithStatus(jetStream);

                    }
                }
                else
                {
                    //console.writeline("no files found.");
                }

            }
        }
    }