namespace MigrationFactory.O365Groups.ExportUsers.Utilities
{
    using MigrationFactory.O365Groups.ExportUsers.Entity;
    using System;
    using System.Collections.Generic;
    using Microsoft.SharePoint.Client;
    using OfficeDevPnP.Core;
    using System.Configuration;
    using MigrationFactory.O365Groups.ModernPage.Utilities;
    using MigrationFactory.O365Groups.ExportUsers.Utilities;
    using System.Threading.Tasks;

    class WebsUtility
    {
        /// <summary>
        /// Gets or sets the modern web.
        /// </summary>
        /// <value>The modern web.</value>
        //public WebSite WebSite { get; set; }
        /// <summary>
        /// The modern webs list
        /// </summary>
        public List<WebSite> WebSiteList = null;
        /// <summary>
        /// Gets or sets the retry count.
        /// </summary>
        /// <value>The retry count.</value>
        public int RetryCount { get; set; }
        /// <summary>
        /// Gets or sets the delay.
        /// </summary>
        /// <value>The delay.</value>
        public int Delay { get; set; }
        /// <summary>
        /// Gets or sets the user agent.
        /// </summary>
        /// <value>The user agent.</value>
        public string UserAgent { get; set; }

        public WebsUtility(int retryCount, int delay, string userAgent)
        {
            WebSiteList = new List<WebSite>();
            RetryCount = retryCount;
            Delay = delay;
            UserAgent = userAgent;
        }

        public (List<string>, List<string>) ExportUsersList()
        {
            List<string> completedSitesList = new List<string>(); 
            List<string> errorSitesList = new List<string>();

            Parallel.ForEach(WebSiteList.ToArray(), (website) =>
            {
                FillWebSiteContext(website);

                if (website.ClientContext != null)
                {
                    (string completedSite, string errorSite) = website.ExportUsers();

                    if(!string.IsNullOrEmpty(completedSite)) completedSitesList.Add(completedSite);
                    if (!string.IsNullOrEmpty(errorSite)) errorSitesList.Add(errorSite);
                }
                else
                    errorSitesList.Add(website.SourceSiteUrl);
            });

            //foreach (var website in WebSiteList.ToArray())
            //{
            //    FillWebSiteContext(website);

            //    if(website.ClientContext != null)
            //        website.ExportUsers();
            //}

            return (completedSitesList, errorSitesList);
        }

        private void FillWebSiteContext(WebSite webSite)
        {
            ConsoleOperations.WriteToConsole("Generating Context for " + webSite.SourceSiteUrl, ConsoleColor.White);

            if (webSite != null)
            {
                var authManager = new AuthenticationManager();

                using (var sourceContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(webSite.SourceSiteUrl, webSite.SourceUserName, webSite.SourcePassword))
                {

                    sourceContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                    {
                        e.WebRequestExecutor.WebRequest.UserAgent = UserAgent;
                    };

                    var rootWeb = sourceContext.Web;
                    try
                    {
                        sourceContext.Load(rootWeb);
                        //sourceContext.ExecuteQueryRetry();
                        sourceContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                        webSite.ClientContext = sourceContext;
                        webSite.Web = rootWeb;
                        WebSiteList.Add(webSite);

                        ConsoleOperations.WriteToConsole("Created context for " + webSite.SourceSiteUrl, ConsoleColor.White);
                    }
                    catch (Exception ex)
                    {
                        //Logger.LogError("Problem in RecursSubWebs for" + WebSite.SourceSiteUrl + Environment.NewLine + ex.Message);
                        ConsoleOperations.WriteToConsole("Problem in FillWebSiteContext for: " + webSite.SourceSiteUrl + Environment.NewLine + "Error Message: " + ex.Message, ConsoleColor.Red);
                    }
                }
            }
        }



    }
}
