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
    using System;
    using System.Linq;
    using MigrationFactory.O365Groups.Logging;
    using MigrationFactory.O365Groups.ModernPage.Utilities;

    /// <summary>
    /// Class RepostPage.
    /// Implements the <see cref="MigrationFactory.O365Groups.ModernPage.ModernPage" />
    /// </summary>
    /// <seealso cref="MigrationFactory.O365Groups.ModernPage.ModernPage" />
    public class RepostPage: ModernPage
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="RepostPage"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="retryCount">The retry count.</param>
        /// <param name="delay">The delay.</param>
        public RepostPage(IAsyncLogger logger, int retryCount, int delay) : base(logger, retryCount, delay)
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
            ListItem repostPageListItem = null;
            Logger.Log("Into CreatePageWithWebParts for " + pageName);

            if (targetWeb != null)
            {
                try
                {
                    var sitePagesList = targetWeb.Lists.GetByTitle(Constants.ModernPageLibrary);
                    var listContentTypes = sitePagesList.ContentTypes;

                    targetContext.Load(targetWeb);
                    targetContext.Load(listContentTypes, ctypes => ctypes.Include(c => c.Name, c => c.Id));
                    targetContext.Load(sitePagesList);
                    targetContext.Load(sitePagesList.RootFolder);
                    targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                    //var pageTitle = "RepostPage.aspx";
                    //sitePagesList.RootFolder.Files.AddTemplateFile(sitePagesList.RootFolder.ServerRelativeUrl + "/" + pageName, TemplateFileType.ClientSidePage);
                    var repostPage = targetWeb.AddClientSidePage(pageName, true);
                    if (repostPage != null)
                    {
                        repostPageListItem = repostPage.PageListItem;

                        repostPageListItem["ContentTypeId"] = listContentTypes.FirstOrDefault(ct => ct.Name == Constants.RepostPageContentType).Id;//sourceItem["ContentTypeId"];//"0x0101009D1CB255DA76424F860D91F20E6C4118002A50BFCFB7614729B56886FADA02339B00874A802FBA36B64BAB7A47514EAAB232";
                        repostPageListItem["PageLayoutType"] = sourceItem["PageLayoutType"];
                        repostPageListItem["PromotedState"] = sourceItem["PromotedState"]; //"2";

                        repostPageListItem["Title"] = sourceItem.Client_Title;
                        //repostPageListItem["ClientSideApplicationId"] = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec";

                        repostPageListItem["_OriginalSourceSiteId"] = targetContext.Site.Id;
                        repostPageListItem["_OriginalSourceWebId"] = targetWeb.Id;
                        repostPageListItem["_OriginalSourceListId"] = sitePagesList.Id;
                        repostPageListItem["_OriginalSourceItemId"] = repostPageListItem["UniqueId"].ToString();

                        repostPageListItem["_OriginalSourceUrl"] = sourceItem["_OriginalSourceUrl"];
                        repostPageListItem["Description"] = sourceItem["Description"];//"Repost Page Description";
                        repostPageListItem["BannerImageUrl"] = sourceItem["BannerImageUrl"];
                        repostPageListItem["FirstPublishedDate"] = sourceItem["FirstPublishedDate"];

                        repostPageListItem.Update();

                        targetContext.Load(repostPageListItem, i => i.DisplayName);
                        targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                        repostPage.PromoteAsNewsArticle();                        
                        
                        if (repostPageListItem != null && sourceItem.HasUniqueRoleAssignments)
                            ManagePermissions(sourceContext, targetContext, sourceItem, repostPageListItem);

                        UpdateSystemFields(targetContext, repostPageListItem, sourceItem);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogError("Error in CreatePageWithWebParts for RepostPage " + pageName + Environment.NewLine + ex.Message);
                    ConsoleOperations.WriteToConsole("Exception in migrating the page: " + pageName + " on " + sourceContext.Web.Url, ConsoleColor.Red);
                    //throw ex;
                }
            }

            return repostPageListItem;
        }
    }
}
