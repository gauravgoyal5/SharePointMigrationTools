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
    //using OfficeDevPnP.Core.Framework.Provisioning.Model;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security;
    using MigrationFactory.O365Groups.Model;
    using MigrationFactory.O365Groups.ModernPage.Entity;
    using MigrationFactory.O365Groups.Logging;
    using MigrationFactory.O365Groups.Factory;
    using MigrationFactory.O365Groups.ModernPage.Utilities;
    /// <summary>
    /// Class Program.
    /// </summary>
    class Program
    {
        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        static void Main(string[] args)
        {
            IAsyncLogger logger = O365GroupsFactory.GetLogger(new string[] { Constants.AppSettings.LoggerInstanceKey }) as IAsyncLogger;
            var siteMapFileName = ConfigurationManager.AppSettings[Constants.AppSettings.SiteMapKey];
            var siteMapSheetName = ConfigurationManager.AppSettings[Constants.AppSettings.SiteMapSheetKey];
            var userMappingFileName = ConfigurationManager.AppSettings[Constants.AppSettings.UserMappingKey];
            var userMappingSheetName = ConfigurationManager.AppSettings[Constants.AppSettings.UserMappingSheetKey];
            var retryCount = int.Parse(ConfigurationManager.AppSettings[Constants.AppSettings.RetryCountKey]);
            var delay = int.Parse(ConfigurationManager.AppSettings[Constants.AppSettings.DelayKey]);
            var userAgent = ConfigurationManager.AppSettings[Constants.AppSettings.UserAgentKey];

            logger.Log("Reading Files...");
            var csvOperation = new CSVOperations();
            var siteMapDetails = csvOperation.ReadFile(Constants.SiteMapSheet, siteMapFileName, siteMapSheetName).Cast<SiteMapReport>().ToList();
            SecurityPrincipal.UserMappingList = csvOperation.ReadFile(Constants.UserMappingSheet, userMappingFileName, userMappingSheetName).Cast<UserMappingReport>().ToList();

            if (siteMapDetails != null)
            {             
                SecureString sourcePassword = GetSecureString(Constants.PasswordMessageSource);
                SecureString targetPassword = GetSecureString(Constants.PasswordMessageTarget);
                Object lockObj = new Object();

                //Stopwatch watch = new Stopwatch();
                //watch.Start();

                List<ModernWeb> websList = new List<ModernWeb>();
                logger.Log("Processing read sites from CSV");

                foreach (var siteMap in siteMapDetails)
                {
                    logger.Log("Reading " + siteMap.SourceSiteUrl);


                    var modernWeb = new ModernWeb
                    {
                        SourceSiteUrl = siteMap.SourceSiteUrl,
                        TargetSiteUrl = siteMap.TargetSiteUrl,
                        SourceUserName = siteMap.SourceUser,
                        TargetUserName = siteMap.TargetUser,
                        SourcePassword = sourcePassword,
                        TargetPassword = targetPassword
                    };

                    WebsOperation websOperation = new WebsOperation(logger, retryCount, delay, userAgent);
                    websOperation.ModernWeb = modernWeb;
                    websOperation.MigrateWebs();
                }

                //watch.Stop();
                logger.Log("Processing complete");
                Console.WriteLine("Processing Complete");
                //Console.WriteLine("Elapsed Time " + watch.Elapsed.ToString());
            }
            
            Console.ReadLine();
        }



        #region Classic Export/Import Webpart Code

        //private static void ImportWebParts(ClientContext targetContext, Web web, Dictionary<string, List<WebPartEntity>> pageWebPartDict)
        //{
        //    var pagesLib = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + "/SitePages");
        //    targetContext.Load(pagesLib);
        //    targetContext.Load(pagesLib.Files);
        //    targetContext.ExecuteQuery();

        //    //File webPartPage = null;

        //    foreach (var pageItem in pageWebPartDict.Keys)
        //    {
        //        var page = targetContext.Web.AddClientSidePage(pageItem, true);

        //        var pageFile = page.Context.Web.GetFileByServerRelativeUrl($"{pagesLib.ServerRelativeUrl}/{pageItem}");
        //        page.Context.Web.Context.Load(pageFile, f => f.ListItemAllFields, f => f.Exists);
        //        page.Context.Web.Context.ExecuteQueryRetry();
        //        //webPartPage = pageItem;
        //        //targetContext.Load(webPartPage);
        //        //targetContext.Load(webPartPage.ListItemAllFields);
        //        //targetContext.ExecuteQuery();

        //        //string wikiField = (string)webPartPage.ListItemAllFields["WikiField"];

        //        LimitedWebPartManager limitedWebPartManager = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);

        //        foreach (var webPartEntity in pageWebPartDict[pageItem])
        //        {
        //            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(webPartEntity.WebPartXml);
        //            WebPartDefinition wpdNew = limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, "wpz", Convert.ToInt32(webPartEntity.WebPartZone));
        //            targetContext.Load(wpdNew); 
        //        }
        //        targetContext.ExecuteQuery();
        //    }            
        //}

        //private static Dictionary<string, List<WebPartEntity>> GetWebPartsToAdd(ClientContext sourceContext, Web web)
        //{
        //    Dictionary<string, List<WebPartEntity>> pageWebPartDict = new Dictionary<string, List<WebPartEntity>>();
        //    List<WebPartEntity> webPartList = new List<WebPartEntity>();

        //    var pagesLib = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + "/SitePages");
        //    sourceContext.Load(pagesLib.Files);
        //    sourceContext.ExecuteQuery();

        //    //Access webparts in each page and perform operation
        //    foreach (var page in pagesLib.Files)
        //    {
        //        sourceContext.Load(page);
        //        sourceContext.Load(page.ListItemAllFields);
        //        sourceContext.ExecuteQuery();

        //        LimitedWebPartManager wpMgr = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
        //        var webPartsDfnCollection = wpMgr.WebParts;
        //        sourceContext.Load(webPartsDfnCollection);
        //        sourceContext.ExecuteQuery();

        //        foreach(var webPartDefn in webPartsDfnCollection)
        //        {
        //            var webPart = webPartDefn.WebPart;
        //            sourceContext.Load(webPart);
        //            sourceContext.ExecuteQuery();

        //            if(webPart.ExportMode != WebPartExportMode.None)
        //            {
        //                Guid webPartId = webPartDefn.Id;
        //                ClientResult<string> webPartXml = wpMgr.ExportWebPart(webPartId);
        //                sourceContext.ExecuteQuery();

        //                WebPartEntity webPartEntity = new WebPartEntity();
        //                webPartEntity.WebPartXml = webPartXml.Value;
        //                webPartEntity.WebPartZone = webPart.ZoneIndex.ToString();
        //                webPartEntity.WebPartTitle = webPart.Title;
        //                webPartList.Add(webPartEntity);
        //            }                    

        //        }
        //        pageWebPartDict.Add(page.Name, webPartList);


        //    }

        //    return pageWebPartDict;
        //}

        #endregion

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
                Console.Write(String.Format("{0}: ", label));

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
