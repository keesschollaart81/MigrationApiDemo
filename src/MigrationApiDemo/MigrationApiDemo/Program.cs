using System;

namespace MigrationApiDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var migrationApiDemo = new MigrationApiDemo();

            // Step 1, create some test files and upload them to a Azure Storage Container
            migrationApiDemo.ProvisionTestFiles();

            // Step 2, create a Migration Package with the job Manifest
            migrationApiDemo.CreateMigrationPackage();

            // Step 3, tell SharePoint where to find the files and start the migration job
            var jobId = migrationApiDemo.StartMigrationJob();

            // Step 4, wait for job to complete, persist logs
            migrationApiDemo.MonitorMigrationApiQueue(jobId).Wait();

            Console.ReadLine();
        }
    }
}