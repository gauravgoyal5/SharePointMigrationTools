namespace MigrationFactory.O365Groups.ExportUsers.Entity
{
    using Microsoft.SharePoint.Client;
    using MigrationFactory.O365Groups.ExportUsers.Utilities;
    using System;
    using System.Security;
    using System.Collections.Generic;
    using System.Linq;
    using System.Configuration;
    using MigrationFactory.O365Groups.Model;
    using System.Text.RegularExpressions;
    using MigrationFactory.O365Groups.ModernPage.Utilities;

    /// <summary>
    /// Class WebSite.
    /// </summary>
    public class WebSite
    {
        /// <summary>
        /// Gets or sets the source site URL.
        /// </summary>
        /// <value>The source site URL.</value>
        public string SourceSiteUrl { get; set; }

        /// <summary>
        /// Gets or sets the name of the source user.
        /// </summary>
        /// <value>The name of the source user.</value>
        public string SourceUserName { get; set; }

        /// <summary>
        /// Gets or sets the source password.
        /// </summary>
        /// <value>The source password.</value>
        public SecureString SourcePassword { get; set; }

        /// <summary>
        /// Gets or sets the client context.
        /// </summary>
        /// <value>The client context.</value>
        public ClientContext ClientContext { get; set; }
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        public Web Web { get; set; }
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

        //public WebSite(int retryCount, int delay, string userAgent)
        //{
        //    RetryCount = retryCount;
        //    Delay = delay;
        //    UserAgent = userAgent;
        //}

        /// <summary>
        /// Exports the users for the current website.
        /// </summary>
        public (string, string) ExportUsers()
        {
            ConsoleOperations.WriteToConsole("Exporting users for website: " + SourceSiteUrl, ConsoleColor.White);

            var domainToSearch = ConfigurationManager.AppSettings[Constants.AppSettings.DomainToSearchKey];
            var exportReportFileName = ConfigurationManager.AppSettings[Constants.AppSettings.UserExportSiteMapReportKey];
            var batchSize = ConfigurationManager.AppSettings[Constants.AppSettings.BatchSizeKey];
            var csvOperation = new CSVOperations();
            string completedSite = string.Empty;
            string erroredSite = string.Empty;

            try
            {
                var userInformationList = Web.SiteUserInfoList;
                var camlQuery = CamlQuery.CreateAllItemsQuery(int.Parse(batchSize), new string[] { "FieldValuesAsText" });
                ClientContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                {
                    e.WebRequestExecutor.WebRequest.UserAgent = UserAgent;
                };

                List<ListItem> items = new List<ListItem>();
                do
                {
                    ListItemCollection listItemCollection = userInformationList.GetItems(camlQuery);
                    ClientContext.Load(listItemCollection);
                    //ClientContext.ExecuteQuery();
                    ClientContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);


                    items.AddRange(listItemCollection);
                    camlQuery.ListItemCollectionPosition = listItemCollection.ListItemCollectionPosition;

                } while (camlQuery.ListItemCollectionPosition != null);

                ConsoleOperations.WriteToConsole($"Extracted all users for {SourceSiteUrl}", ConsoleColor.Green);
                IEnumerable<User> domainUsers = GetDesiredDomainUsers(items, domainToSearch);
                               
                ConsoleOperations.WriteToConsole($"Writing Users List for {SourceSiteUrl}. Total users found are: " + domainUsers.Count(), ConsoleColor.Yellow);

                bool isSuccess = csvOperation.WriteCsv(domainUsers, exportReportFileName);
                if (isSuccess)
                    completedSite = SourceSiteUrl;
                else
                    erroredSite = SourceSiteUrl;

            }
            catch (Exception ex)
            {
                ConsoleOperations.WriteToConsole($"Error in Exporting users {ex.Message}", ConsoleColor.Red);
                    erroredSite = SourceSiteUrl;                
            }

            return (completedSite, erroredSite);
        }

        private static IEnumerable<User> GetDesiredDomainUsers(List<ListItem> items, string domainToSearch)
        {
            var domainUsers = items.Where(item =>
            {
                string stringToSearch = item.FieldValues["Name"].ToString(); 
                string pattern = $@"\.?({domainToSearch})#?";

                Match match = Regex.Match(stringToSearch, pattern, RegexOptions.IgnoreCase);

                if (match.Success)
                {
                    return true;
                }
                else
                    return false;
            })
            .Select(item =>
            new User
            {
                UPN = item.FieldValues["Name"].ToString(),
                Name = item.FieldValues["Title"].ToString()
            });

            //var domainUsers = items.Where(item => item.FieldValues["Name"].ToString().Contains(domainToSearch))
            //                      .Select(item =>
            //                        new
            //                        {
            //                            UPN = item.FieldValues["Name"].ToString(),
            //                            Name = item.FieldValues["Title"].ToString()
            //                        });



            //var domainUsers = items.Select(item =>
            //                        new
            //                        {
            //                            UPN = item.FieldValues["Name"].ToString(),
            //                            Name = item.FieldValues["Title"].ToString()
            //                        });

            return domainUsers;
        }

        //private (RoleAssignmentCollection roleAssignments, GroupCollection groups) GetRolesAndGroups()
        //{
        //    var roleAssignments = Web.RoleAssignments;
        //    var webGroups = roleAssignments.Groups;
        //    ClientContext.Load(webGroups);
        //    ClientContext.Load(roleAssignments);
        //    ClientContext.ExecuteQuery();

        //    return (roleAssignments, webGroups);
        //}
    }
}
