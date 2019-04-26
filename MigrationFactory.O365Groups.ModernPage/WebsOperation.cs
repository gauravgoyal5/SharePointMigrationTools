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
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using MigrationFactory.O365Groups.ModernPage.Utilities;
    using MigrationFactory.O365Groups.ModernPage.Entity;
    using MigrationFactory.O365Groups.Logging;
    using System.Threading.Tasks;

    /// <summary>
    /// Class WebsOperation.
    /// </summary>
    public class WebsOperation
    {
        /// <summary>
        /// Gets or sets the modern web.
        /// </summary>
        /// <value>The modern web.</value>
        public ModernWeb ModernWeb { get; set; }
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

        /// <summary>
        /// The modern webs list
        /// </summary>
        List<ModernWeb> ModernWebsList = null;


        /// <summary>
        /// The logger
        /// </summary>
        IAsyncLogger Logger = null;
        /// <summary>
        /// Initializes a new instance of the <see cref="WebsOperation" /> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="retryCount">The retry count.</param>
        /// <param name="delay">The delay.</param>
        /// <param name="userAgent">The user agent.</param>
        public WebsOperation(IAsyncLogger logger, int retryCount, int delay, string userAgent)
        {            
            ModernWebsList = new List<ModernWeb>();
            Logger = logger;
            RetryCount = int.Parse(ConfigurationManager.AppSettings[Constants.AppSettings.RetryCountKey]);
            Delay = int.Parse(ConfigurationManager.AppSettings[Constants.AppSettings.DelayKey]);
            UserAgent = ConfigurationManager.AppSettings[Constants.AppSettings.UserAgentKey];
        }

        /// <summary>
        /// Migrates the webs.
        /// </summary>
        public void MigrateWebs()
        {
            RecursSubWebs();

            //For parallel operation
            Parallel.ForEach(ModernWebsList, (modernWeb) =>
            {
                //Got the root web. Get page items collection
                var sourcePageItemsCollection = GetPageItemsCollection(modernWeb);

                //Use this PageItemsCollection to create pages and add webparts
                ConnectTarget(modernWeb, sourcePageItemsCollection);
            });

            ////For serial operation/non-parallel
            //foreach (var modernWeb in ModernWebsList)
            //{
            //    //Got the root web. Get page items collection
            //    var sourcePageItemsCollection = GetPageItemsCollection(modernWeb);

            //    //Use this PageItemsCollection to create pages and add webparts
            //    var targetItem = ConnectTarget(modernWeb, sourcePageItemsCollection);
            //}

        }


        /// <summary>
        /// Recurses the sub webs of a site collection and generate the list of webs for further processing.
        /// </summary>
        public void RecursSubWebs()
        {
            Logger.Log("Reading webs for " + ModernWeb.SourceSiteUrl);
            Console.WriteLine("Reading webs for " + ModernWeb.SourceSiteUrl);
            
            if (ModernWeb != null)
            {
                var authManagerSource = new AuthenticationManager();
                                
                using (var sourceContext = authManagerSource.GetSharePointOnlineAuthenticatedContextTenant(ModernWeb.SourceSiteUrl, ModernWeb.SourceUserName, ModernWeb.SourcePassword))
                {
                    Logger.Log("Created context for " + ModernWeb.SourceSiteUrl);

                    sourceContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                    {
                        e.WebRequestExecutor.WebRequest.UserAgent = UserAgent;
                    };

                    var rootWeb = sourceContext.Web;
                    try
                    {
                        sourceContext.Load(rootWeb,
                                        website => website.RoleDefinitions,
                                        website => website.WebTemplate,
                                        website => website.Webs,
                                        website => website.Title,
                                        website => website.Url);
                        //sourceContext.ExecuteQueryRetry();
                        sourceContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                        ModernWeb.ClientContext = sourceContext;
                        ModernWeb.Web = rootWeb;

                        ModernWebsList.Add(ModernWeb);

                        //Get all the webs for the current root site
                        Logger.Log("Calling GetAllSubWebs for " + ModernWeb.SourceSiteUrl);
                        var websList = GetAllSubWebs(sourceContext, rootWeb);

                        websList?.ForEach(web =>
                       {
                           ModernWeb = web;
                           RecursSubWebs();
                       });
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError("Problem in RecursSubWebs for" + ModernWeb.SourceSiteUrl + Environment.NewLine + ex.Message);
                        ConsoleOperations.WriteToConsole("Problem in RecursSubWebs for: " + ModernWeb.SourceSiteUrl + Environment.NewLine + "Error Message: " + ex.Message, ConsoleColor.Red);
                    }
                } 
            }
        }

        /// <summary>
        /// Gets all sub webs for a given web.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="rootWeb">The root web.</param>
        /// <returns>List&lt;ModernWeb&gt;.</returns>
        public List<ModernWeb> GetAllSubWebs(ClientContext context, Web rootWeb)
        {
            Logger.Log("Enterned GetAllSubWebs for " + rootWeb.Url);

            List<ModernWeb> modernWebsList = new List<ModernWeb>(); var userAgent = ConfigurationManager.AppSettings[Constants.AppSettings.UserAgentKey];
            try
            {
                var websCollection = rootWeb.Webs;
                context.Load(websCollection, webs => webs.Include(web => web.Url));
                context.ExecuteQueryWithIncrementalRetry(RetryCount, Delay); 

                foreach (var web in websCollection)
                {
                    string sourceNewpath = web.Url;//sourceDict["SiteUrl"] + web.ServerRelativeUrl;
                    string subsiteRelativeUrl = sourceNewpath.ToUpperInvariant().Replace(ModernWeb.SourceSiteUrl.ToUpperInvariant(), String.Empty);
                    //SourceDictionary["SiteUrl"] = sourceNewpath;

                    string targetNewpath = ModernWeb.TargetSiteUrl + subsiteRelativeUrl;//targetDict["SiteUrl"] + web.ServerRelativeUrl;
                                                                                        //TargetDictionary["SiteUrl"] = targetNewpath;

                    modernWebsList.Add(new ModernWeb
                    {
                        SourceSiteUrl = sourceNewpath,
                        TargetSiteUrl = targetNewpath,
                        SourceUserName = ModernWeb.SourceUserName,
                        SourcePassword = ModernWeb.SourcePassword,
                        TargetUserName = ModernWeb.TargetUserName,
                        TargetPassword = ModernWeb.TargetPassword
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("Error in GetAllWebs" + Environment.NewLine + ex.Message);
                throw ex;
            }

            return modernWebsList;

        }

        /// <summary>
        /// Gets the page items collection.
        /// </summary>
        /// <param name="modernWeb">The modern web.</param>
        /// <returns>ListItemCollection.</returns>
        public ListItemCollection GetPageItemsCollection(ModernWeb modernWeb)
        {
            Logger.Log("Enterned GetPageItemsCollection for " + modernWeb.SourceSiteUrl);
            ListItemCollection items = null;
            if (modernWeb != null)
            {
                var sourceContext = modernWeb.ClientContext;
                var rootWeb = modernWeb.Web;
                var webTemplate = modernWeb.Web.WebTemplate;

                //if (webTemplate == Constants.WebTemplate)
                //{
                try
                {
                    var listTitle = Constants.ModernPageLibrary;
                    var list = rootWeb.Lists.GetByTitle(listTitle);
                    var migrationType = ConfigurationManager.AppSettings[Constants.AppSettings.MigrationTypeKey];

                    //Check if we only want Modern Page to work with
                    if (MigrationType.All == (MigrationType) Enum.Parse(typeof(MigrationType),migrationType))
                    {
                        items = list.GetItems(CamlQuery.CreateAllItemsQuery()); 
                    }
                    else
                    {
                        var query = new CamlQuery();
                        query.ViewXml = Constants.ModernPageQueryString;
                        items = list.GetItems(query);
                    }
                    modernWeb.ClientContext.Load(items, icol => icol.Include(
                            i => i[Constants.FirstPublishedDate],
                            i => i[Constants.BannerImageUrl],
                            i => i[Constants.ContentTypeId],
                            i => i[Constants.Description],
                            i => i[Constants.OriginalSourceUrl],
                            i => i[Constants.PromotedState],
                            i => i[Constants.PageLayoutType],
                            i => i[Constants.Author],
                            i => i[Constants.Editor],
                            i => i[Constants.Created],
                            i => i[Constants.Modified],
                            i => i[Constants.FileLeafRef],
                            i => i[Constants.FileRef], 
                            i => i[Constants.PageLayoutType], 
                            i => i[Constants.ClientSideApplicationId], 
                            i => i[Constants.LayoutWebpartsContent],
                            i => i[Constants.WikiPageContentControl], 
                            i => i[Constants.ModernPageContentControl],
                            i => i.ContentType, 
                            i => i.DisplayName,
                            i => i.ParentList, 
                            i => i.Client_Title, 
                            i => i.HasUniqueRoleAssignments, 
                            i => i.RoleAssignments,
                            i => i.RoleAssignments.Include(
                                ra => ra.Member,
                                    ra => ra.Member.LoginName,
                                    ra => ra.Member.PrincipalType,
                                    ra => ra.RoleDefinitionBindings,
                                    ra => ra.RoleDefinitionBindings.Include(rd => rd.RoleTypeKind))));
                    sourceContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);
                }
                catch (Exception ex)
                {
                    Logger.LogError("Error in GetPageItemsCollection" + Environment.NewLine + ex.Message);
                    throw ex;
                }
                //} 
            }

            return items;
        }

        /// <summary>
        /// Connects the target.
        /// </summary>
        /// <param name="modernWeb">The modern web.</param>
        /// <param name="sourcePageItemsCollection">The source page items collection.</param>
        private ListItem ConnectTarget(ModernWeb modernWeb, ListItemCollection sourcePageItemsCollection)
        {
            Logger.Log("Entered ConnectTarget for " + modernWeb.TargetSiteUrl);
            ListItem targetItem = null;
            if (modernWeb != null)
            {
                try
                {
                    var sourceContext = modernWeb.ClientContext;
                    var authManagerTarget = new AuthenticationManager();

                    using (var targetContext = authManagerTarget.GetSharePointOnlineAuthenticatedContextTenant(modernWeb.TargetSiteUrl, modernWeb.TargetUserName, modernWeb.TargetPassword))
                    {
                        sourceContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                        {
                            e.WebRequestExecutor.WebRequest.UserAgent = UserAgent;
                        };

                        var targetSite = targetContext.Site;
                        var targetWeb = targetContext.Web;

                        targetContext.Load(targetSite);
                        targetContext.Load(targetWeb,
                                w => w.SiteGroups,
                                w => w.SiteGroups.Include(sg => sg.LoginName),
                                w => w.Title,
                                w => w.RoleDefinitions,
                                w => w.WebTemplate);
                        targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                        //var webTemplate = web.WebTemplate;
                        //if (webTemplate == Constants.WebTemplate)
                        //{
                        targetItem = CreatePage(sourcePageItemsCollection, sourceContext, targetContext, targetWeb);

                        //}
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogError("Error in Connecting to target site " + modernWeb.TargetSiteUrl + Environment.NewLine + ex.Message);
                    throw ex;
                }
            }

            return targetItem;
        }

        /// <summary>
        /// Creates the page.
        /// </summary>
        /// <param name="sourcePageItemsCollection">The source page items collection.</param>
        /// <param name="sourceContext">The source context.</param>
        /// <param name="targetContext">The target context.</param>
        /// <param name="targetWeb">The target web.</param>
        private ListItem CreatePage(ListItemCollection sourcePageItemsCollection, ClientContext sourceContext, ClientContext targetContext, Web targetWeb)
        {
            ListItem targetItem = null;
            if (sourcePageItemsCollection != null && sourcePageItemsCollection.Count > 0)
            {
                foreach (var sourceItem in sourcePageItemsCollection)
                {
                    var fileRef = sourceItem[Constants.FileRef].ToString();                    
                    var pageName = fileRef.Substring(fileRef.LastIndexOf('/') + 1);

                    ConsoleOperations.WriteToConsole("Processing page: " + fileRef, ConsoleColor.Yellow);
                    Logger.Log("Processing page: " + fileRef);

                    


                    try
                    {
                        switch (sourceItem.ContentType.Name)
                        {
                            case Constants.ModernPageContentType:
                                //Sometimes Wiki Pages have 'Site Page' content Type. Additional check for that
                                //TODO: Use Dependency Injection/Factory Model
                                if (sourceItem[Constants.WikiPageContentControl] == null)
                                {                                    
                                    SitePage modernPage = new SitePage(Logger, RetryCount, Delay);
                                    targetItem = modernPage.CreatePageWithWebParts(sourceContext, targetContext, targetWeb, sourceItem, pageName);
                                }
                                else
                                {
                                    WikiPage wikiPage = new WikiPage(Logger, RetryCount, Delay);
                                    targetItem = wikiPage.CreatePageWithWebParts(sourceContext, targetContext, targetWeb, sourceItem, pageName);
                                }

                                ConsoleOperations.WriteToConsole("Processed page: " + fileRef, ConsoleColor.Green);
                                Logger.Log("Processed page: " + fileRef);
                                break;
                            case Constants.WikiPageContentType:
                                //TODO: If a page is present then avoid adding duplicate webparts  
                                WikiPage wikPage = new WikiPage(Logger, RetryCount, Delay);
                                targetItem = wikPage.CreatePageWithWebParts(sourceContext, targetContext, targetWeb, sourceItem, pageName);

                                ConsoleOperations.WriteToConsole("Processed page: " + fileRef, ConsoleColor.Green);
                                Logger.Log("Processed page: " + fileRef);

                                break;

                            case Constants.RepostPageContentType:
                                RepostPage repostPage = new RepostPage(Logger, RetryCount, Delay);
                                targetItem = repostPage.CreatePageWithWebParts(sourceContext, targetContext, targetWeb, sourceItem, pageName);

                                ConsoleOperations.WriteToConsole("Processed page: " + fileRef, ConsoleColor.Green);
                                Logger.Log("Processed page: " + fileRef);

                                break;

                            case Constants.WebPartPageContentType:
                                //TODO - Multiple webpart Zone
                                WebPartPage webPartPage = new WebPartPage(Logger, RetryCount, Delay);
                                webPartPage.CreatePageWithWebParts(sourceContext, targetContext, targetWeb, sourceItem, pageName);

                                ConsoleOperations.WriteToConsole("Processed page: " + fileRef, ConsoleColor.Green);
                                Logger.Log("Processed page: " + fileRef);
                                break;

                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError("Exception in WebsOperation for page content type identification. " + ex.Message);
                    }
                } 
            }

            return targetItem;
        }



        //private ListItem CreateWikiPageWithWebParts(ClientContext sourceContext, ClientContext targetContext, Web web, ListItem item, string pageName)
        //{
        //    ListItem targetItem;
        //    var wikiPageRelativeUrl = web.EnsureWikiPage(item.ParentList.Title, pageName);//.AddWikiPage(Constants.ModernPageLibrary, pageName);
        //    var wikiPage = web.GetFileByServerRelativeUrl(web.ServerRelativeUrl + "/" + wikiPageRelativeUrl);

        //    targetItem = wikiPage.ListItemAllFields;
        //    targetContext.Load(wikiPage);
        //    targetContext.Load(targetItem);
        //    targetContext.ExecuteQuery();

        //    var webpartList = WebPartOperation.ExportWebPart(sourceContext, item.File);
        //    WebPartOperation.ImportWebParts(targetContext, web, webpartList, wikiPage);

        //    AddContent(targetContext, web, targetItem, item, Constants.WikiPageContentControl);
        //    return targetItem;
        //}

        //private void AddContent(ClientContext targetContext, Web web, ListItem targetItem, ListItem sourceItem, string pageControlId)
        //{
        //    targetItem[pageControlId] = sourceItem[pageControlId];
        //    targetItem.Update();
        //    targetContext.ExecuteQuery();
        //}


    }
}
