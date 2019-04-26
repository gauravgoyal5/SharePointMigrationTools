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
    using OfficeDevPnP.Core;
    using OfficeDevPnP.Core.Pages;
    using System;
    using MigrationFactory.O365Groups.Logging;
    using MigrationFactory.O365Groups.ModernPage.Utilities;

    /// <summary>
    /// Class SitePage.
    /// Implements the <see cref="MigrationFactory.O365Groups.ModernPage.ModernPage" />
    /// </summary>
    /// <seealso cref="MigrationFactory.O365Groups.ModernPage.ModernPage" />
    public class SitePage : ModernPage
    {
        //IAsyncLogger Logger = null;
        //public int RetryCount { get; set; }
        //public int Delay { get; set; }
        /// <summary>
        /// Initializes a new instance of the <see cref="SitePage"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="retryCount">The retry count.</param>
        /// <param name="delay">The delay.</param>
        public SitePage(IAsyncLogger logger, int retryCount, int delay) : base(logger, retryCount, delay)
        {
            Logger = logger;
            RetryCount = retryCount;
            Delay = delay;
        }
        /// <summary>
        /// Creates a modern page with web parts.
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

            try
            {                
                var newClientSidePage = targetWeb.AddClientSidePage(pageName, true);

                try
                {
                    var sourcePage = ClientSidePage.Load(sourceContext, pageName);

                    if (sourcePage.PageHeader.ImageServerRelativeUrl != null)
                    {
                        newClientSidePage.PageHeader.ImageServerRelativeUrl = sourcePage.PageHeader.ImageServerRelativeUrl.ToLower().Replace(sourceContext.Web.Url, targetContext.Web.Url);
                    }

                    newClientSidePage.PageTitle = sourcePage.PageTitle;
                    newClientSidePage.LayoutType = sourcePage.LayoutType;

                    if (sourcePage.LayoutType == ClientSidePageLayoutType.Home)
                    {
                        newClientSidePage.RemovePageHeader();
                    }

                    if (sourcePage.CommentsDisabled)
                    {
                        newClientSidePage.DisableComments();
                    }
                }
                catch (Exception ex)
                {
                    ConsoleOperations.WriteToConsole("Problem in fetching source page: " + pageName + " at: " + sourceContext.Web.Url + ex.Message, ConsoleColor.Red);
                    Logger.LogError("Problem in fetching source page: " + pageName + " at: " + sourceContext.Web.Url + ex.Message);
                }

                if (!string.IsNullOrEmpty(Convert.ToString(sourceItem[Constants.PromotedState])))
                {
                    if (Convert.ToInt32(sourceItem[Constants.PromotedState]) == 2)
                    {
                        newClientSidePage.PromoteAsNewsArticle();
                    }
                }
                newClientSidePage.Save();

                targetItem = newClientSidePage.PageListItem;

                targetContext.Load(targetItem, i => i.DisplayName);
                targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                AddContent(targetContext, targetItem, sourceItem, Constants.ModernPageContentControl);

                if (targetItem != null && sourceItem.HasUniqueRoleAssignments)
                    ManagePermissions(sourceContext, targetContext, sourceItem, targetItem);

                UpdateSystemFields(targetContext, targetItem, sourceItem);

                //newClientSidePage.Publish();
            }
            catch (Exception ex)
            {
                Logger.LogError("Error in CreatePageWithWebParts for ModernPage " + pageName + Environment.NewLine + ex.Message);
                ConsoleOperations.WriteToConsole("Exception in migrating the page: " + pageName + " on " + sourceContext.Web.Url, ConsoleColor.Red);
                //throw ex;
            }
            //AddControls(sourceContext, newClientSidePage, pageName);
            

            //var webpartList = WebPartOperation.ExportWebPart(sourceContext, sourceItem.File);
            //WebPartOperation.ImportWebParts(targetContext, web, webpartList, newClientSidePage.PageListItem.File);
            
            return targetItem;
        }

        /// <summary>
        /// Adds the controls.
        /// </summary>
        /// <param name="sourceContext">The source context.</param>
        /// <param name="newModernPage">The new modern page.</param>
        /// <param name="pageName">Name of the page.</param>
        private void AddControls(ClientContext sourceContext, ClientSidePage newModernPage, string pageName)
        {
            try
            {
                var sourcePage = ClientSidePage.Load(sourceContext, pageName); //newModernPage.Controls;
                var controlsCollection = sourcePage.Controls;

                foreach (var control in controlsCollection)
                {
                    newModernPage.AddControl(control);                    
                    //var webparttoadd = newmodernpage.instantiatedefaultwebpart(defaultclientsidewebparts.)
                }
                newModernPage.Save();
            }
            catch (Exception ex)
            {
                ConsoleOperations.WriteToConsole("Problem with adding controls to page", ConsoleColor.Yellow); 
            }
        }
    }
}
