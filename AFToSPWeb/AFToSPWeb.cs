using System;
using System.Security;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;

namespace AFToSPWeb
{
    public static class AFToSPWeb
    {
        [FunctionName("AFToSPWeb")]
        public static void Run([TimerTrigger("0 0 9-17 26-27 * *")]TimerInfo myTimer, TraceWriter log)
        {
            try
            {
                log.Info($"C# Timer trigger function executed at: {DateTime.Now}");
                string siteUrl = "https://northwindgrp1425.sharepoint.com/";
                string userName = GetEnvironmentVariable("SPUser");
                string password = GetEnvironmentVariable("SPPwd");
                SecureString pwd = new SecureString();
                foreach (char pwdChar in password)
                    pwd.AppendChar(pwdChar);

                string SPData = string.Empty;
                SharePointOnlineCredentials creds = new SharePointOnlineCredentials(userName, pwd);

                using (ClientContext clientContext = new ClientContext(siteUrl))
                {
                    clientContext.Credentials = creds;

                    Web web = clientContext.Web;

                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    SPData += "Title: " + web.Title + ", Last Modified: " + web.LastItemModifiedDate.ToString();

                    log.Info(SPData);

                }

            }
            catch (Exception ex)
            {
                log.Info(ex.ToString());
            }
        }

        public static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
    }
}
