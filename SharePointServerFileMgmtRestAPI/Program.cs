using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.IO;

namespace SharePointServerFileMgmtRestAPI
{
    class Program : IDisposable
    {
        const int TWO_HUNDRED_FIFTY_MB = 250 * 1024 * 1024; // 250 MB
        const int CHUNK_SIZE_BYTES     = TWO_HUNDRED_FIFTY_MB; // max 250 mb

        static HttpClient _httpClient = null;

        static readonly string _baseAddress   = "https://hnsc.2019.contoso.lab";
        static readonly string _largeFilePath = @"C:\_temp\large.txt";
        static readonly string _smallFilePath = @"C:\_temp\small.txt";


        static readonly string _relativeUrl = "/sites/teamsite";
        static readonly string _folderPath  = "/sites/teamsite/Shared%20Documents";

        // GetFolderByServerRelativeUrl, GetFileByServerRelativeUrl and GetByUrlOrAddStub will allow support for SharePoint Server 2016

        // single file paths
        static readonly string _uploadPath = $"{_relativeUrl}/_api/web/GetFolderByServerRelativeUrl('{_folderPath}')/Files/Add(overwrite=true, url='{{0}}')";

        // chunked file paths
        static readonly string _startUploadPath = $"{_relativeUrl}/_api/web/GetFolderByServerRelativeUrl('{_folderPath}')/Files/GetByUrlOrAddStub('{{0}}')/StartUpload(uploadId=guid'{{1}}')";
        static readonly string _continueUploadPath = $"{_relativeUrl}/_api/web/GetFileByServerRelativeUrl('{_folderPath}/{{0}}')/ContinueUpload(uploadId=guid'{{1}}',fileOffset={{2}})";
        static readonly string _finishUploadPath   = $"{_relativeUrl}/_api/web/GetFileByServerRelativeUrl('{_folderPath}/{{0}}')/FinishUpload(uploadId=guid'{{1}}',fileOffset={{2}})";

        static readonly string _downloadPath = $"{_relativeUrl}/_api/web/GetFileByServerRelativeUrl('{{0}}')/$value";

        static readonly string _recyclePath = $"{{0}}/recycle";

        static readonly string _deletePath = $"{{0}}/deleteObject";

        static readonly string _getListItemPath = $"{_relativeUrl}/_api/web/GetFileByServerRelativeUrl('{{0}}')/ListItemAllFields";

        // form digest path
        static readonly string _contextPath = "/sites/teamsite/_api/contextinfo";

        /* Fiddler trace
       
            /sites/teamsite/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/Files/AddStubUsingPath(DecodedUrl=@a2)/StartUpload(uploadId=@a3)?
                @a1='/sites/teamsite/Shared Documents'&
                @a2='file1.txt'&
                @a3=guid'eb7b0d9e-b57b-42e7-aa95-f9b9a508c528'

            /sites/teamsite/_api/web/GetFileByServerRelativePath(DecodedUrl= @a1)/ContinueUpload(uploadId= @a2, fileOffset= @a3)?
                @a1='/sites/teamsite/Shared Documents/file1.txt'
                &@a2=guid'eb7b0d9e-b57b-42e7-aa95-f9b9a508c528'&
                @a3=0

            /sites/teamsite/_api/web/GetFileByServerRelativePath(DecodedUrl= @a1)/FinishUpload(uploadId= @a2, fileOffset= @a3)?
                @a1='/sites/teamsite/Shared Documents/file1.txt'&
                @a2=guid'eb7b0d9e-b57b-42e7-aa95-f9b9a508c528'&
                @a3=786432000	
        */

        static void Main(string[] args)
        {
            Init();

            UploadFile(_largeFilePath);

            UploadFile(_smallFilePath);
            
            DownloadFile($"{_folderPath}/small.txt", @"C:\_temp\small.txt");

            RecycleFile($"{_folderPath}/small.txt");

            DeleteFile($"{_folderPath}/example.txt");
        }

        static void Init()
        {
            if( null == _httpClient)
            {
                var httpClientHandler = new HttpClientHandler()
                {
                    UseDefaultCredentials = true
                    //Credentials = new System.Net.NetworkCredential("contoso\\johndoe", "secret password")
                };

                _httpClient = new HttpClient(httpClientHandler)
                {
                    BaseAddress = new System.Uri(_baseAddress),
                };

                _httpClient.DefaultRequestHeaders.Add("accept", "application/json;odata=verbose");
                _httpClient.DefaultRequestHeaders.Add("X-RequestDigest", GetFormDigestValue());
            }
        }

        static void DownloadFile(string webRelativeFileUrl, string path)
        {
            // start download
            var bytes = _httpClient
                .GetByteArrayAsync(string.Format(_downloadPath, webRelativeFileUrl))
                .GetAwaiter()
                .GetResult();

            // write bytes to disk
            System.IO.File.WriteAllBytes(path, bytes);
        }

        static void DeleteFile(string webRelativeFileUrl)
        {
            // get list item details based on the file url
            var response = _httpClient
                .GetAsync(string.Format(_getListItemPath, webRelativeFileUrl))
                .GetAwaiter()
                .GetResult();

            // parse json response
            dynamic parsed = JsonConvert.DeserializeObject<dynamic>(response.Content.ReadAsStringAsync().Result);

            // pull the uri value uri=https://hnsc.2019.contoso.lab/sites/teamsite/_api/Web/Lists(guid'b73a2191-362d-47b6-9291-9ae9f43aaf00')/Items(30)
            var itemUri = (string)parsed.d.__metadata.uri;

            // execute recycle request
            _ = _httpClient
                .PostAsync(string.Format(_deletePath, itemUri), null)
                .GetAwaiter()
                .GetResult()
                .EnsureSuccessStatusCode();
        }

        static void RecycleFile(string webRelativeFileUrl)
        {
            // get list item details based on the file url
            var response = _httpClient
                .GetAsync(string.Format(_getListItemPath, webRelativeFileUrl))
                .GetAwaiter()
                .GetResult();

            // parse json response
            dynamic parsed = JsonConvert.DeserializeObject<dynamic>(response.Content.ReadAsStringAsync().Result);

            // pull the uri value uri=https://hnsc.2019.contoso.lab/sites/teamsite/_api/Web/Lists(guid'b73a2191-362d-47b6-9291-9ae9f43aaf00')/Items(30)
            var itemUri = (string)parsed.d.__metadata.uri;

            // execute recycle request
            _ = _httpClient
                .PostAsync(string.Format(_recyclePath, itemUri), null)
                .GetAwaiter()
                .GetResult()
                .EnsureSuccessStatusCode();
        }

        static void UploadFile(string path)
        {
            var fi = new System.IO.FileInfo(path);

            if (fi.Length > TWO_HUNDRED_FIFTY_MB)
            {
                UploadLargeFile(path);
            }
            else
            {
                UploadSmallFile(path);
            }
        }

        static void UploadSmallFile(string path)
        {
            var fi = new System.IO.FileInfo(path);

            if( fi.Length > TWO_HUNDRED_FIFTY_MB)
            {
                UploadLargeFile(path);
                return;
            }

            string uniqueFileName = GetTimeStampedFileName(fi.Name);

            Console.WriteLine($"Uploading {fi.FullName} as {uniqueFileName}");

            _ = _httpClient
                    .PostAsync(string.Format(_uploadPath, uniqueFileName), new ByteArrayContent(File.ReadAllBytes(fi.FullName)))
                    .GetAwaiter()
                    .GetResult()
                    .EnsureSuccessStatusCode();
        }

        static void UploadLargeFile(string path)
        {
            var fi = new FileInfo(path);

            string uniqueFileName = GetTimeStampedFileName(fi.Name);

            Console.WriteLine($"Uploading {fi.FullName} as {uniqueFileName}");

            // open the file
            using (var fileStream = File.OpenRead(fi.FullName))
            {
                fileStream.Position = 0;

                // generate a unique guid for the upload session
                string uploadId = Guid.NewGuid().ToString();

                long offset   = 0;
                int bytesRead = 0;
                
                var bytes = new byte[CHUNK_SIZE_BYTES];

                do
                {
                    if (fileStream.Position == 0)
                    {
                        // start chunked upload
                        _ = _httpClient
                                .PostAsync(string.Format(_startUploadPath, uniqueFileName, uploadId), null)
                                .GetAwaiter()
                                .GetResult()
                                .EnsureSuccessStatusCode();
                    }
                    else if (fileStream.Position < fileStream.Length)
                    {
                        // continue chunked upload
                        var response = _httpClient
                                .PostAsync(string.Format(_continueUploadPath, uniqueFileName, uploadId, offset), new ByteArrayContent(bytes, 0, bytesRead))
                                .GetAwaiter()
                                .GetResult()
                                .EnsureSuccessStatusCode();

                        // read the response to find the offset for the next chunk
                        dynamic parsed = JsonConvert.DeserializeObject<dynamic>(response.Content.ReadAsStringAsync().Result);

                        offset = (int)parsed.d.ContinueUpload;

                    }
                    else if (fileStream.Position == fileStream.Length)
                    {
                        // finish chunked upload
                        _ = _httpClient
                                .PostAsync(string.Format(_finishUploadPath, uniqueFileName, uploadId, offset), new ByteArrayContent(bytes, 0, bytesRead))
                                .GetAwaiter()
                                .GetResult()
                                .EnsureSuccessStatusCode();
                    }

                    bytesRead = fileStream.Read(bytes, 0, bytes.Length);
                }
                while (bytesRead > 0);
            }
        }

        static string GetFormDigestValue()
        {
            var response = _httpClient
                .PostAsync(_contextPath, null)
                .GetAwaiter()
                .GetResult()
                .EnsureSuccessStatusCode();

            // parse json response
            dynamic parsed = JsonConvert.DeserializeObject<dynamic>( response.Content.ReadAsStringAsync().Result );

            // return FormDigestValue
            return parsed.d.GetContextWebInformation.FormDigestValue; // example response: 0xF11B65FE8814D7931F89FADE490B7C43FC562E5488D9D6D9D5D96556030B98E8513C6D5F8B3ADB628A178A4B49E3D95CBFDE4AAAB1C43610F714AE7C5C7D303A,17 Sep 2021 15:11:22 -0000
        }

        static string GetTimeStampedFileName(string fileName)
        {
            // turns foo.txt into foo_132766403889693739.txt
            return $"{System.IO.Path.GetFileNameWithoutExtension(fileName)}_{DateTime.Now.ToFileTime()}{System.IO.Path.GetExtension(fileName)}";
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}
