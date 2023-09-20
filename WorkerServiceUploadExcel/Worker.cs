using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Net;

namespace WorkerServiceUploadExcel
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly IHostApplicationLifetime _hostApplicationLifetime;
        private IConfiguration _configuration;
        private HttpClient httpClient;
        private string _ftpServerUrl, _ftpPort, _ftpUsername, _ftpPassword, _remoteFolderPath;

        public Worker(ILogger<Worker> logger, IHostApplicationLifetime hostApplicationLifetime, IConfiguration configuration)
        {
            _logger = logger;
            _hostApplicationLifetime = hostApplicationLifetime;
            _configuration = configuration;
        }

        public override Task StartAsync(CancellationToken cancellationToken)
        {
            httpClient = new HttpClient();
            _ftpServerUrl = _configuration["FTPSetting:FtpHost"];
            _ftpPort = _configuration["FTPSetting:FtpPort"];
            _ftpUsername = _configuration["FTPSetting:FtpUser"];
            _ftpPassword = _configuration["FTPSetting:FtpPass"];
            _remoteFolderPath = _configuration["FTPSetting:FtpFolder"];
            _logger.LogInformation("================ Service started at: {time}", DateTimeOffset.Now);
            return base.StartAsync(cancellationToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            string remoteFilePath = _remoteFolderPath + "/example.xlsx";
            string ftpServerUrl = _ftpServerUrl + ":" +_ftpPort;

            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("======= Worker running at: {time}", DateTimeOffset.Now);

                var stream = new MemoryStream();
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(stream))
                {
                    var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                    // Ghi dữ liệu vào các ô của worksheet
                    workSheet.Cells[1, 1].Value = "Hello";
                    workSheet.Cells[1, 2].Value = "World";
                    workSheet.Cells[2, 1].Value = "This";
                    workSheet.Cells[2, 2].Value = "is";
                    workSheet.Cells[3, 1].Value = "an";
                    workSheet.Cells[3, 2].Value = "example";
                    package.Save();
                }
                stream.Position = 0;

                // Upload file lên FTP server

                using (WebClient client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(_ftpUsername, _ftpPassword);
                    client.UploadData(ftpServerUrl + remoteFilePath, stream.ToArray());
                }

                Console.WriteLine("File uploaded successfully!");


                await Task.Delay(1000, stoppingToken);

                // Dừng service
                _hostApplicationLifetime.StopApplication();
                await Task.Delay(1000);
            }
        }

        public bool UploadFileViaFtp(string filePath, string ftpHost, string ftpFolder, string ftpUser, string ftpPass, int chunkSize)
        {
            var success = false;
            FileStream fs = null;
            Stream rs = null;

            try
            {
                string uploadFileName = new FileInfo(filePath).Name;
                string uploadUrl = ftpHost + "/" + ftpFolder;
                //fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                FileInfo fileInf = new FileInfo(filePath);
                fs = fileInf.OpenRead();
                string ftpUrl = string.Format("{0}/{1}", uploadUrl, uploadFileName);
                FtpWebRequest requestObj = FtpWebRequest.Create(ftpUrl) as FtpWebRequest;
                requestObj.KeepAlive = false;
                requestObj.UseBinary = false;
                requestObj.Method = WebRequestMethods.Ftp.UploadFile;
                requestObj.Credentials = new NetworkCredential(ftpUser, ftpPass);
                ////rs = requestObj.GetRequestStream();
                //int chunkSize = 10485760;//10Mb
                //int.TryParse(_config["AppSettings:ChunkSize"], out chunkSize);
                //if (chunkSize == null || chunkSize <= 0)
                //{
                //    chunkSize = 10485760;//10Mb
                //}
                Stream strm = requestObj.GetRequestStream();
                requestObj.ContentLength = fileInf.Length;
                int contentLen;
                byte[] buff = new byte[chunkSize];
                contentLen = fs.Read(buff, 0, chunkSize);

                while (contentLen != 0)
                {
                    // Write Content from the file stream to the FTP Upload
                    // Stream
                    strm.Write(buff, 0, contentLen);
                    contentLen = fs.Read(buff, 0, chunkSize);
                }

                // Close the file stream and the Request Stream
                strm.Close();
                //rs.Flush();
                fs.Close();
                //byte[] buffer = new byte[chunkSize];
                //int read = 0;
                //while ((read = fs.Read(buffer, 0, buffer.Length)) != 0)
                //{
                //    rs.Write(buffer, 0, read);
                //}
                //rs.Flush();

                success = true;
            }
            catch (Exception ex)
            {
                _logger.LogError("UploadFileViaFtp", ex.Message, ex.InnerException);
                success = false;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }

                if (rs != null)
                {
                    rs.Close();
                    rs.Dispose();
                }
            }
            return success;
        }
    }
}