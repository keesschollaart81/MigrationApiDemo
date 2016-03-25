using System;

namespace MigrationApiDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var migrationApiDemo = new MigrationApiDemo();

            // Step 1, Create and upload some test-files to Azure Blob Storage
            migrationApiDemo.ProvisionTestFiles();

            // Step 2, Create and upload Manifest Package to Azure Blob Storage
            migrationApiDemo.CreateAndUploadMigrationPackage();

            // Step 3, Start the Migration Job using SharePoint Online Clientside Object Model (CSOM)
            var jobId = migrationApiDemo.StartMigrationJob();

            // Step 4, Monitor the Reporting Queue, persist messages/logs and wait for the job to complete
            migrationApiDemo.MonitorMigrationApiQueue(jobId).Wait();

            Console.ReadLine();
        }
    }
}