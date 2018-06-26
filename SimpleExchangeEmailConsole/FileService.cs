using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace SimpleExchangeEmailConsole
{
    public static class FileService
    {
        public  static string FileExtenton { get; set; }

        private static readonly string Folder =
            Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\Reports\";

        private static readonly string ArchiveFolder
            = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\Archive\";

        public static List<string> GetAllFiles()
        {
            Logger.Info("Getting files from Reports folder");
            return Directory.GetFiles(Folder, FileExtenton).ToList();
        }

        public static void MoveFile(string sourceFile)
        {
            if (!Directory.Exists(ArchiveFolder))
            {
                Logger.Info("Created Archive directory");

                Directory.CreateDirectory(ArchiveFolder);
            }

            var targetPath = ArchiveFolder + $"{DateTime.Now:MMM yy}";
            if (!Directory.Exists(targetPath))
            {
                Logger.Info($"Created Archive directory for {DateTime.Now:MMMM} month");

                Directory.CreateDirectory(targetPath);
            }

            var fileName = Path.GetFileName(sourceFile);
            var destFile = Path.Combine(targetPath, fileName);
            File.Copy(sourceFile, destFile, true);
            File.Delete(sourceFile);
            Logger.Info($"File {fileName} moved to archive");
        }
    }
}
