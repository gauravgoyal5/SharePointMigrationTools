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
    using System.Threading.Tasks;
    using MigrationFactory.O365Groups.ModernPage.Utilities;
    using MigrationFactory.O365Groups.Logging;
    using Newtonsoft.Json.Linq;
    using System.Web;
    //using Microsoft.SharePoint.Client.Utilities;

    /// <summary>
    /// Class ModernPage.
    /// </summary>
    public abstract class ModernPage
    {
        /// <summary>
        /// Gets or sets the logger.
        /// </summary>
        /// <value>The logger.</value>
        public IAsyncLogger Logger { get; set; }
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
        /// Initializes a new instance of the <see cref="ModernPage"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="retryCount">The retry count.</param>
        /// <param name="delay">The delay.</param>
        public ModernPage(IAsyncLogger logger, int retryCount, int delay)
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
        /// <param name="web">The web.</param>
        /// <param name="sourceItem">The source item.</param>
        /// <param name="pageName">Name of the page.</param>
        /// <returns>ListItem.</returns>
        public abstract ListItem CreatePageWithWebParts(ClientContext sourceContext, ClientContext targetContext, Web web, ListItem sourceItem, string pageName);

        /// <summary>
        /// Adds the content.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <param name="targetItem">The target item.</param>
        /// <param name="sourceItem">The source item.</param>
        /// <param name="pageControlId">The page control identifier.</param>
        protected void AddContent(ClientContext targetContext, ListItem targetItem, ListItem sourceItem, string pageControlId)
        {
            Logger.Log("Into AddContent for " + sourceItem.DisplayName);

            if (sourceItem != null)
            {
                try
                {
                    targetItem[pageControlId] = sourceItem[pageControlId];

                    targetItem[Constants.PageLayoutType] = sourceItem[Constants.PageLayoutType];
                    targetItem[Constants.ClientSideApplicationId] = sourceItem[Constants.ClientSideApplicationId];
                    targetItem[Constants.LayoutWebpartsContent] = sourceItem[Constants.LayoutWebpartsContent];

                    targetItem.UpdateOverwriteVersion();
                    targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);
                }
                catch (Exception ex)
                {
                    Logger.LogError("Error in AddContent for " + sourceItem.DisplayName + Environment.NewLine + ex.Message);
                    //throw ex;
                } 
            }
        }

        /// <summary>
        /// Updates the system fields.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <param name="targetItem">The target item.</param>
        /// <param name="sourceItem">The source item.</param>
        protected ListItem UpdateSystemFields(ClientContext targetContext, ListItem targetItem, ListItem sourceItem)
        {
            Logger.Log("Into UpdateSystemFields for " + sourceItem.DisplayName);

            if (sourceItem != null && targetItem != null)
            {
                try
                {
                    var securityPrincipal = new SecurityPrincipal(Logger, RetryCount, Delay);
                    FieldUserValue author = securityPrincipal.GetTargetUserFromUserMapping(targetContext, ((FieldUserValue)sourceItem[Constants.Author]));
                    FieldUserValue editor = securityPrincipal.GetTargetUserFromUserMapping(targetContext, ((FieldUserValue)sourceItem[Constants.Editor]));

                    targetItem[Constants.Author] = author; //authorUpn.Substring(authorUpn.LastIndexOf('|') + 1)
                    targetItem[Constants.Editor] = editor; //((FieldUserValue)sourceItem["Editor"]).LookupValue
                    targetItem[Constants.Created] = (DateTime)sourceItem[Constants.Created];
                    targetItem[Constants.Modified] = (DateTime)sourceItem[Constants.Modified];

                    //targetItem.Update();
                    targetItem.UpdateOverwriteVersion();
                    //targetContext.Load(targetItem);
                    //targetContext.ExecuteQuery();
                    targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);
                }
                catch (Exception ex)
                {
                    Logger.LogError("Error in AddContent for " + sourceItem.DisplayName + Environment.NewLine + ex.Message);
                    //throw ex;
                }
            }

            return targetItem;
        }

        /// <summary>
        /// Manages the permissions.
        /// </summary>
        /// <param name="sourceContext">The source context.</param>
        /// <param name="targetContext">The target context.</param>
        /// <param name="sourceItem">The source item.</param>
        /// <param name="targetItem">The target item.</param>
        protected void ManagePermissions(ClientContext sourceContext, ClientContext targetContext, ListItem sourceItem, ListItem targetItem)
        {
            if (targetContext != null && sourceItem != null && targetItem != null)
            {
                Logger.Log("Into ManagePermissions for " + sourceItem.DisplayName);

                try
                {
                    var targetWeb = targetContext.Web;
                    var sourceItemRoleAssignmentCollection = sourceItem.RoleAssignments;

                    targetContext.Load(targetItem, ti => ti.RoleAssignments);
                    targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

                    //var sourceRoleDefinitionCollection = sourceContext.Web.RoleDefinitions;
                    targetItem.BreakRoleInheritance(false, true);

                    foreach (var sourceItemRoleAssignment in sourceItemRoleAssignmentCollection)
                    {
                        ManageSecurityPrincipal(targetContext, targetItem, targetWeb, sourceItemRoleAssignment);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogError("Error in Manage Permissions for " + sourceItem.DisplayName + Environment.NewLine + ex.Message);
                    //throw ex;
                }
            }
        }

        /// <summary>
        /// Manages the security principal.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <param name="targetItem">The target item.</param>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="sourceItemRoleAssignment">The source item role assignment.</param>
        private void ManageSecurityPrincipal(ClientContext targetContext, ListItem targetItem, Web targetWeb, RoleAssignment sourceItemRoleAssignment)
        {
            if (sourceItemRoleAssignment != null)
            {
                Task log = Logger.Log("Into ManageSecurityPrincipal for " + sourceItemRoleAssignment.Member);

                try
                {
                    var raMember = sourceItemRoleAssignment.Member;
                    var principalType = raMember.PrincipalType;
                    var roleType = sourceItemRoleAssignment.RoleDefinitionBindings[0].RoleTypeKind;

                    SecurityPrincipal securityPrincipal = new SecurityPrincipal(Logger, RetryCount, Delay);

                    var task = Task.Run(() =>
                    {
                        Principal principal = securityPrincipal.GetSecurityPrincipal(targetContext, raMember, principalType);
                        if (principal != null)
                        {
                            securityPrincipal.SetSecurityPrincipal(targetContext, targetItem, targetWeb, roleType, principal);
                        }
                        else
                            Logger.LogWarning("Security principal not found for " + raMember.LoginName + 
                                Environment.NewLine +
                                "Item affected: " + 
                                targetItem.DisplayName);
                    });

                    task.Wait();
                    if (task.Status == TaskStatus.RanToCompletion)
                        Logger.Log("Permission migration completed for " + targetItem.DisplayName);
                    else
                        Logger.LogError("Permission migration did not complete properly for " + targetItem.DisplayName + 
                            Environment.NewLine +
                            "Task Status" +
                            task.Status.ToString());
                }
                catch (Exception ex)
                {
                    log = Logger.LogError("Error in ManageSecurityPrincipal for " + sourceItemRoleAssignment.Member + 
                        Environment.NewLine + 
                        ex.Message);
                    //throw ex;
                }
            }
        }

        


    }
}
