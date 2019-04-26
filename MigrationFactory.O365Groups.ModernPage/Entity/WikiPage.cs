// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.ModernPage
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
namespace MigrationFactory.O365Groups.ModernPage
{
    using Microsoft.SharePoint.Client;
    //using OfficeDevPnP.Core.Framework.Provisioning.Model;
    using System;
    using MigrationFactory.O365Groups.Logging;
    using MigrationFactory.O365Groups.ModernPage.Utilities;

    /// <summary>
    /// Class WikiPage.
    /// Implements the <see cref="MigrationFactory.O365Groups.ModernPage.ModernPage" />
    /// </summary>
    /// <seealso cref="MigrationFactory.O365Groups.ModernPage.ModernPage" />
    public class WikiPage: ModernPage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WikiPage"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="retryCount">The retry count.</param>
        /// <param name="delay">The delay.</param>
        public WikiPage(IAsyncLogger logger, int retryCount, int delay) : base(logger, retryCount, delay)
        {
            Logger = logger;
            RetryCount = retryCount;
            Delay = delay;
        }
        /// <summary>
        /// Creates the page with web parts.
        /// </summary>
        /// <param name="sourceContext">The source context.</param>
        /// <param name="targetContext">The target context.</param>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="sourceItem">The source item.</param>
        /// <param name="pageName">Name of the page.</param>
        /// <returns>ListItem.</returns>
        public override ListItem CreatePageWithWebParts(ClientContext sourceContext, ClientContext targetContext, Web targetWeb, ListItem sourceItem, string pageName)
        {
            ListItem targetItem = null;
            Logger.Log("Into CreatePageWithWebParts for " + pageName);

            if (targetWeb != null)
            {
                try
                {
                    var wikiPageRelativeUrl = targetWeb.EnsureWikiPage(sourceItem.ParentList.Title, pageName); //.AddWikiPage(Constants.ModernPageLibrary, pageName);
                    var wikiPage = targetWeb.GetFileByServerRelativeUrl(targetWeb.ServerRelativeUrl + "/" + wikiPageRelativeUrl);

                    targetItem = wikiPage.ListItemAllFields;
                    targetContext.Load(wikiPage);
                    targetContext.Load(targetItem, i => i.DisplayName);
                    targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                    var webPartOps = new WebPartOperation(Logger, RetryCount, Delay);
                    var webpartList = webPartOps.ExportWebPart(sourceContext, sourceItem.File);
                    webPartOps.ImportWebParts(targetContext, targetWeb, webpartList, wikiPage);

                    AddContent(targetContext, targetItem, sourceItem, Constants.WikiPageContentControl);
                    
                    if (targetItem != null && sourceItem.HasUniqueRoleAssignments)
                        ManagePermissions(sourceContext, targetContext, sourceItem, targetItem);

                    UpdateSystemFields(targetContext, targetItem, sourceItem);
                }
                catch (Exception ex)
                {
                    ConsoleOperations.WriteToConsole("Exception in migrating the page: " + pageName + " on " + sourceContext.Web.Url, ConsoleColor.Red);
                    Logger.LogError("Error in CreatePageWithWebParts for WikiPage " + pageName + Environment.NewLine + ex.Message);
                    //throw ex;
                }
            }

            return targetItem;
        }
    }
}
