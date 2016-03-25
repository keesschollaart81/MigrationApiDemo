using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using log4net;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage.Queue;

namespace MigrationApiDemo
{
    public class MigrationApiDemo
    {
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private ICollection<SourceFile> _filesToMigrate;
        private readonly AzureBlob _blobContainingManifestFiles;
        private readonly SharePointMigrationTarget _target;
        private readonly AzureCloudQueue _migrationApiQueue;
        private readonly TestDataProvider _testDataProvider;

        public MigrationApiDemo()
        {
            Log.Debug("Initiaing SharePoint connection.... ");

            _target = new SharePointMigrationTarget();

            Log.Debug("Initiating Storage for test files, manifest en reporting queue");

            _blobContainingManifestFiles = new AzureBlob(
                ConfigurationManager.AppSettings["ManifestBlob.AccountName"],
                ConfigurationManager.AppSettings["ManifestBlob.AccountKey"],
                ConfigurationManager.AppSettings["ManifestBlob.ContainerName"]);

            var testFilesBlob = new AzureBlob(
                ConfigurationManager.AppSettings["SourceFilesBlob.AccountName"],
                ConfigurationManager.AppSettings["SourceFilesBlob.AccountKey"],
                ConfigurationManager.AppSettings["SourceFilesBlob.ContainerName"]);

            _testDataProvider = new TestDataProvider(testFilesBlob);

            _migrationApiQueue = new AzureCloudQueue(
                ConfigurationManager.AppSettings["ReportQueue.AccountName"],
                ConfigurationManager.AppSettings["ReportQueue.AccountKey"],
                ConfigurationManager.AppSettings["ReportQueue.QueueName"]);
        }

        public void ProvisionTestFiles()
        {
            _filesToMigrate = _testDataProvider.ProvisionAndGetFiles();
        }

        public void CreateAndUploadMigrationPackage()
        {
            if (!_filesToMigrate.Any()) throw new Exception("No files to create Migration Package for, run ProvisionTestFiles() first!");

            var manifestPackage = new ManifestPackage(_target);
            var filesInManifestPackage = manifestPackage.GetManifestPackageFiles(_filesToMigrate);

            var blobContainingManifestFiles = _blobContainingManifestFiles;
            blobContainingManifestFiles.RemoveAllFiles();

            foreach (var migrationPackageFile in filesInManifestPackage)
            {
                blobContainingManifestFiles.UploadFile(migrationPackageFile.Filename, migrationPackageFile.Contents);
            }
        }

        /// <returns>Job Id</returns>
        public Guid StartMigrationJob()
        {
            var sourceFileContainerUrl = _testDataProvider.GetBlobUri();
            var manifestContainerUrl = _blobContainingManifestFiles.GetUri(
                SharedAccessBlobPermissions.Read 
                | SharedAccessBlobPermissions.Write 
                | SharedAccessBlobPermissions.List);

            var azureQueueReportUrl = _migrationApiQueue.GetUri(
                SharedAccessQueuePermissions.Read 
                | SharedAccessQueuePermissions.Add 
                | SharedAccessQueuePermissions.Update 
                | SharedAccessQueuePermissions.ProcessMessages);

            return _target.StartMigrationJob(sourceFileContainerUrl, manifestContainerUrl, azureQueueReportUrl);
        }

        private void DownloadAndPersistLogFiles(Guid jobId)
        {
            foreach (var filename in _blobContainingManifestFiles.ListFilenames())
            {
                if (filename.StartsWith($"Import-{jobId}"))
                {
                    Log.Debug($"Downloaded logfile {filename}");
                    File.WriteAllBytes(filename, _blobContainingManifestFiles.DownloadFile(filename));
                }
            }
        }

        public async Task MonitorMigrationApiQueue(Guid jobId)
        {
            while (true)
            {
                var message = await _migrationApiQueue.GetMessageAsync<UpdateMessage>();
                if (message == null)
                {
                    await Task.Delay(TimeSpan.FromSeconds(1));
                    continue;
                }

                switch (message.Event)
                {
                    case "JobEnd":
                        Log.Info($"Migration Job Ended {message.FilesCreated:0.} files created, {message.TotalErrors:0.} errors.!");
                        DownloadAndPersistLogFiles(jobId); // save log files to disk
                        Console.WriteLine("Press ctrl+c to exit");
                        return;
                    case "JobStart":
                        Log.Info("Migration Job Started!");
                        break;
                    case "JobProgress":
                        Log.Debug($"Migration Job in progress, {message.FilesCreated:0.} files created, {message.TotalErrors:0.} errors.");
                        break;
                    case "JobQueued":
                        Log.Info("Migration Job Queued...");
                        break;
                    case "JobWarning":
                        Log.Warn($"Migration Job warning {message.Message}");
                        break;
                    case "JobError":
                        Log.Error($"Migration Job error {message.Message}");
                        break;
                    default:
                        Log.Warn($"Unknown Job Status: {message.Event}, message {message.Message}");
                        break;

                }
            }
        }
    }
}