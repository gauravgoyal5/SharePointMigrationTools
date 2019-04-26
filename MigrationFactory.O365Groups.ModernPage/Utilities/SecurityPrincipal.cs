// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.ModernPage.Utilities
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
namespace MigrationFactory.O365Groups.ModernPage.Utilities
{
    using Microsoft.SharePoint.Client;
    using OfficeDevPnP.Core;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Threading.Tasks;
    using MigrationFactory.O365Groups.Model;
    using Microsoft.SharePoint.Client.Utilities;
    using MigrationFactory.O365Groups.Logging;
    /// <summary>
    /// Class SecurityPrincipal.
    /// </summary>
    public class SecurityPrincipal
    {
        /// <summary>
        /// The logger
        /// </summary>
        IAsyncLogger Logger = null;
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
        /// Gets or sets the user mapping list.
        /// </summary>
        /// <value>The user mapping list.</value>
        public static List<UserMappingReport> UserMappingList { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="SecurityPrincipal" /> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="retryCount">The retry count.</param>
        /// <param name="delay">The delay.</param>
        public SecurityPrincipal(IAsyncLogger logger, int retryCount, int delay)
        {
            Logger = logger;
            RetryCount = retryCount;
            Delay = delay;
        }
        /// <summary>
        /// Gets the security principal.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <param name="raMember">The ra member.</param>
        /// <param name="principalType">Type of the principal.</param>
        /// <returns>Principal.</returns>
        public Principal GetSecurityPrincipal(ClientContext targetContext, Principal raMember, PrincipalType principalType)
        {
            Group group = null;
            Principal principal = null;
            
            try
            {
                switch (principalType)
                {
                    case PrincipalType.SharePointGroup:
                        var groupName = FindGroup(targetContext.Web, raMember.LoginName);
                        group = targetContext.Web.SiteGroups.GetByName(groupName);
                        principal = group;
                        break;
                    case PrincipalType.User:
                        var user = GetTargetUserFromUserMapping(targetContext, raMember.LoginName.Substring(raMember.LoginName.LastIndexOf('|') + 1));//targetWeb.EnsureUser(raMember.LoginName);//FindGroup(targetWeb, raMember.LoginName);
                        principal = user;
                        break;
                    default:
                        break;

                }
            }
            catch (Exception ex)
            {
                Task log = Logger.LogError("Exception in GetSecurityPrincipal " + ex.Message);
                throw ex;
            }            

            return principal;
        }

        /// <summary>
        /// Sets the security principal.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <param name="targetItem">The target item.</param>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="roleType">Type of the role.</param>
        /// <param name="principal">The principal.</param>
        public void SetSecurityPrincipal(ClientContext targetContext, ListItem targetItem, Web targetWeb, RoleType roleType, Principal principal)
        {
            if (principal != null)
            {
                try
                {
                    var role = targetWeb.RoleDefinitions.GetByType(roleType); //RoleType.Contributor/Administrator/Visitor;
                    var roleDefinitionsBindingColl = new RoleDefinitionBindingCollection(targetContext) { role };

                    targetItem.RoleAssignments.Add(principal, roleDefinitionsBindingColl);
                    targetItem.Update();
                    targetContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);
                }
                catch (Exception ex)
                {
                    Logger.LogError("Exception in SetSecurityPrincipal " + ex.Message);
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Gets the target user from user mapping.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <param name="sourceUser">The source user.</param>
        /// <returns>FieldUserValue.</returns>
        public FieldUserValue GetTargetUserFromUserMapping(ClientContext targetContext, FieldUserValue sourceUser)
        {
            FieldUserValue targetUserFieldValue = new FieldUserValue();
            User targetUser = null;
            if (!string.IsNullOrEmpty(sourceUser.Email))
            {
                try
                {
                    var targetUserString = UserMappingList.SingleOrDefault(r => r.SourceUserId == sourceUser.Email).TargetUserId;
                    targetUser = GetUsersDetails(targetContext, targetUserString);
                }
                catch (Microsoft.SharePoint.Client.ServerException serverEx)
                {
                    Logger.LogError("Error in resolving the user. Details: " + serverEx.Message);
                    ConsoleOperations.WriteToConsole($"Email value for the user {sourceUser.Email} not found in the Target.", ConsoleColor.Yellow);
                    targetUser = GetDefaultUser(targetContext);
                }
                catch (Exception ex)
                {
                    Logger.LogError("Error in resolving the user. Details: " + ex.Message);
                    ConsoleOperations.WriteToConsole($"Email value for the user {sourceUser.Email} not found in the Target.", ConsoleColor.Yellow);
                    targetUser = GetDefaultUser(targetContext);
                }
            }
            else if(!string.IsNullOrEmpty(sourceUser.LookupValue))
            {
                try
                {
                    targetUser = GetUsersDetails(targetContext, sourceUser.LookupValue);
                }
                catch (Exception ex)
                {
                    Logger.LogError($"Look up value for the user {sourceUser.LookupValue} not found in the Target" + ex.Message);
                    ConsoleOperations.WriteToConsole($"Look up value for the user {sourceUser.LookupValue} not found in the Target.", ConsoleColor.Yellow);
                    targetUser = GetDefaultUser(targetContext);
                }
            }
            else
            {
                targetUser = GetDefaultUser(targetContext);
            }

            targetUserFieldValue.LookupId = targetUser.Id;
            return targetUserFieldValue;
        }


        /// <summary>
        /// Gets the target user from user mapping file.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <param name="sourceUserName">Name of the source user.</param>
        /// <returns>User.</returns>
        public User GetTargetUserFromUserMapping(ClientContext targetContext, string sourceUserName)
        {
            User targetUser = null;
            if (!string.IsNullOrEmpty(sourceUserName))
            {
                try
                {
                    var targetUserString = UserMappingList.SingleOrDefault(r => r.SourceUserId == sourceUserName).TargetUserId;
                    targetUser = GetUsersDetails(targetContext, targetUserString);
                }
                catch (Microsoft.SharePoint.Client.ServerException serverEx)
                {
                    Logger.LogError("Error in resolving the user. Details: " + serverEx.Message);
                    targetUser = GetDefaultUser(targetContext);
                }
                catch(Exception ex)
                {
                    Logger.LogError("Error in resolving the user. Details: " + ex.Message);
                    targetUser = GetDefaultUser(targetContext);
                }
            }
            else
            {
                targetUser = GetDefaultUser(targetContext);
            }

            return targetUser;
        }

        /// <summary>
        /// Gets the users details.
        /// </summary>
        /// <param name="clientContext">The client context.</param>
        /// <param name="uName">Name of the u.</param>
        /// <returns>User.</returns>
        public User GetUsersDetails(ClientContext clientContext, string uName)
        {
            User newUser = clientContext.Web.EnsureUser(uName);
            clientContext.Load(newUser);
            clientContext.ExecuteQueryWithIncrementalRetry(RetryCount, Delay);

            return newUser;
        }

        /// <summary>
        /// Gets the default user.
        /// </summary>
        /// <param name="targetContext">The target context.</param>
        /// <returns>User.</returns>
        private User GetDefaultUser(ClientContext targetContext)
        {
            var targetUser = GetUsersDetails(targetContext, ConfigurationManager.AppSettings[O365Groups.ModernPage.Constants.AppSettings.DefaultUserKey]);//targetContext.Web.EnsureUser(ConfigurationManager.AppSettings[O365Groups.ModernPage.Constants.AppSettings.DefaultUserKey]);
            return targetUser;
        }

        /// <summary>
        /// Finds the group.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="loginName">Name of the login.</param>
        /// <returns>System.String.</returns>
        private static string FindGroup(Web targetWeb, string loginName)
        {
            var groupName = GetGroup(targetWeb, loginName);

            if (groupName != string.Empty)
                return groupName;
            else
            {
                var loginNameWords = loginName.Split(' ');
                var newGroupName = targetWeb.SiteGroups.FirstOrDefault(g => g.LoginName.Contains(loginNameWords.Last())).LoginName;

                return GetGroup(targetWeb, newGroupName);
            }
        }

        /// <summary>
        /// Gets the group.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="loginName">Name of the login.</param>
        /// <returns>System.String.</returns>
        private static string GetGroup(Web targetWeb, string loginName)
        {
            string groupName = string.Empty;

            var doesGroupExists = targetWeb.GroupExists(loginName);
            if (doesGroupExists)
                groupName = loginName;

            return groupName;
        }
    }
}
