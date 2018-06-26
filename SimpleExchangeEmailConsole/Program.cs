using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace SimpleExchangeEmailConsole
{
    public class Program
    {
        private static readonly NameValueCollection AppSettings = ConfigurationManager.AppSettings;
    
        static void Main(string[] args)
        {
            Logger.IsLoggingEnabed = bool.Parse(AppSettings["Logging"]);
            Logger.Info("App started", true);
            FileService.FileExtenton = AppSettings["FileExtention"];

            Exchange.CreateService(AppSettings);

            var files = FileService.GetAllFiles();
            var undelivered = new List<string>();
            var sentCount = 0;
            foreach (var file in files)
            {
                if (Exchange.SendEmail(file))
                {
                    sentCount++;
                }
                else
                {
                    undelivered.Add(Path.GetFileNameWithoutExtension(file));
                }
            }
            Logger.Info($"Delivered {sentCount} attachments", true);
            if (undelivered.Count > 0)
            {
                Logger.Info($"Undelivered to: ", true);
                foreach (var user in undelivered)
                {
                    Logger.Info(user, true);
                }
            }
            Logger.Info("Press enter to exit", true, true);
        }
    }
}
