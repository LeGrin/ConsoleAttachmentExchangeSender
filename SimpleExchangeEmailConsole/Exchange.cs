using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace SimpleExchangeEmailConsole
{
    public static class Exchange
    {
        private static NameValueCollection AppSettings { get; set; }
        public static ExchangeService Service { get; set; }
        public static void CreateService(NameValueCollection appSettings)
        {
            try
            {
                AppSettings = appSettings;
                Logger.Info("Getting user credentials");
                var username = AppSettings["UserName"];
                var userdomain = AppSettings["UserDomain"];
                var login = username + "@" + userdomain;
                var password = AppSettings["Password"];

                Enum.TryParse(AppSettings["ExchangeVersion"], out ExchangeVersion exchangeVersion);
                Service = new ExchangeService(exchangeVersion)
                {
                    Credentials = new WebCredentials(login, password)
                };

                if (bool.Parse(AppSettings["Trace"]))
                {
                    Service.TraceEnabled = true;
                    Service.TraceFlags = TraceFlags.All;
                }

                Logger.Info("Starting domain server auto discovery...");
                Service.AutodiscoverUrl(login, RedirectionUrlValidationCallback);
                Logger.Info("Domain server found");

                Logger.Info("Requested server version " + Service.RequestedServerVersion);
                Logger.Info("Server url " + Service.Url);
                Logger.Info("User Name of sender person: " + GetUserDisplayName(username));

            }
            catch (Exception e)
            {
                Logger.Exception("####################SOME ISSUE ENCOUNTERED. COPY TEXT BEETWEN THIS LINES######################");
                Logger.Exception(e.Message);
                Logger.Exception("####################SOME ISSUE ENCOUNTERED. COPY TEXT BEETWEN THIS LINES######################", true);
                throw;
            }
        }

        public static bool SendEmail(string fileName)
        {
            var username = Path.GetFileNameWithoutExtension(fileName);
            if (GetUserDisplayName(username) == string.Empty) return false;
            try
            {
                Logger.Info($"Preparing email for {username}");
                EmailMessage email = new EmailMessage(Service);
                email.ToRecipients.Add(username + "@" + AppSettings["EmailDomain"]);
                email.Subject = AppSettings["Subject"];
                email.Body = new MessageBody(AppSettings["Body"]);
                email.Attachments.AddFileAttachment(fileName);

                email.Send();
                FileService.MoveFile(fileName);
                Logger.Info($"Salary statement sent to {username}");
                return true;
            }
            catch (Exception e)
            {
                Logger.Exception("FAILED TO SEND MESSAGE TO " + username);
                Logger.Exception(
                    "####################SOME ISSUE ENCOUNTERED. COPY TEXT BEETWEN THIS LINES######################");
                Logger.Exception(e.Message);
                Logger.Exception(
                    "####################SOME ISSUE ENCOUNTERED. COPY TEXT BEETWEN THIS LINES######################", true);
                return false;
            }
        }

        public static string GetUserDisplayName(string username)
        {
            try
            {
                return Service.ResolveName(username, ResolveNameSearchLocation.ContactsThenDirectory, true)[0].Contact
                    .DisplayName;
            }
            catch
            {
                Logger.Exception("User name not found in Active Directory: " + username);
                return string.Empty;
            }

        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

    }
}
