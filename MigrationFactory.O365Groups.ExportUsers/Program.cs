// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.ExportUsers
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************

namespace MigrationFactory.O365Groups.ExportUsers
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security;
    using MigrationFactory.O365Groups.Model;
    using MigrationFactory.O365Groups.Logging;
    using MigrationFactory.O365Groups.Factory;
    using MigrationFactory.O365Groups.ExportUsers.Entity;
    using MigrationFactory.O365Groups.ExportUsers.Utilities;

    class Program
    {
        static void Main(string[] args)
        {
            IAsyncLogger logger = O365GroupsFactory.GetLogger(new string[] { Constants.AppSettings.LoggerInstanceKey }) as IAsyncLogger;
            var siteMapFileName = ConfigurationManager.AppSettings[Constants.AppSettings.UserExportSiteMapKey];
            var siteMapSheetName = ConfigurationManager.AppSettings[Constants.AppSettings.SiteMapSheetKey];
            var retryCount = int.Parse(ConfigurationManager.AppSettings[Constants.AppSettings.RetryCountKey]);
            var delay = int.Parse(ConfigurationManager.AppSettings[Constants.AppSettings.DelayKey]);
            var userAgent = ConfigurationManager.AppSettings[Constants.AppSettings.UserAgentKey];
            List<string> completedSitesList = new List<string>();
            List<string> erroredSitesList = new List<string>();

            ConsoleOperations.WriteToConsole("Reading Files...", ConsoleColor.White);
            var csvOperation = new CSVOperations();
            var siteMapDetails = csvOperation.ReadFile(Constants.AppSettings.UserExportSiteMapKey, siteMapFileName, siteMapSheetName).Cast<UserExportSiteMapReport>().ToList();

            if (siteMapDetails != null)
            {
                SecureString sourcePassword = GetSecureString(Constants.PasswordMessageSource);
                //Stopwatch watch = new Stopwatch();
                //watch.Start();

                List<WebSite> websList = new List<WebSite>();
                ConsoleOperations.WriteToConsole("Processing read sites from CSV", ConsoleColor.White);

                foreach (var siteMap in siteMapDetails)
                {
                    logger.Log("Reading " + siteMap.SourceSiteUrl);


                    var website = new WebSite
                    {
                        SourceSiteUrl = siteMap.SourceSiteUrl.Trim(),
                        SourceUserName = siteMap.SourceUser.Trim(),
                        SourcePassword = sourcePassword,
                        Delay = delay,
                        RetryCount = retryCount,
                        UserAgent = userAgent
                    };
                    websList.Add(website);
                }

                try
                {
                    WebsUtility websOperation = new WebsUtility(retryCount, delay, userAgent);
                    websOperation.WebSiteList = websList;
                    (completedSitesList, erroredSitesList) = websOperation.ExportUsersList();
                }
                catch (Exception ex)
                {
                    ConsoleOperations.WriteToConsole("Error ocurred: " + ex.Message, ConsoleColor.Red);
                }

                //watch.Stop();
                logger.Log("Processing complete");
                ConsoleOperations.WriteToConsole($"Processing Complete; Total successful site count: {completedSitesList.Count}. The completed list is as follows:", ConsoleColor.Green);
                completedSitesList.ForEach(site =>
                {
                    ConsoleOperations.WriteToConsole(site, ConsoleColor.DarkCyan);
                });

                ConsoleOperations.WriteToConsole("Total Errored sites: " + erroredSitesList.Count, ConsoleColor.Red);
                if (erroredSitesList.Count>0)
                {
                    ConsoleOperations.WriteToConsole("List: ", ConsoleColor.Red);
                    erroredSitesList.ForEach(site =>
                    {
                        ConsoleOperations.WriteToConsole(site, ConsoleColor.Red);
                    }); 
                }
                
                Console.ReadLine();
                //Console.WriteLine("Elapsed Time " + watch.Elapsed.ToString());
            }
        }

        /// <summary>
        /// Gets password from the user and converts to Secure String
        /// </summary>
        /// <param name="label">The label.</param>
        /// <returns>SecureString.</returns>
        private static SecureString GetSecureString(string label)
        {
            SecureString sStrPwd = new SecureString();
            try
            {
                ConsoleOperations.WriteToConsole(label, ConsoleColor.Yellow);

                for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
                {
                    if (keyInfo.Key == ConsoleKey.Backspace)
                    {
                        if (sStrPwd.Length > 0)
                        {
                            sStrPwd.RemoveAt(sStrPwd.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                    }
                    else if (keyInfo.Key != ConsoleKey.Enter)
                    {
                        Console.Write("*");
                        sStrPwd.AppendChar(keyInfo.KeyChar);
                    }

                }
                Console.WriteLine("");
            }
            catch (Exception e)
            {
                sStrPwd = null;
                Console.WriteLine(e.Message);
            }

            return sStrPwd;
        }
    }
}
