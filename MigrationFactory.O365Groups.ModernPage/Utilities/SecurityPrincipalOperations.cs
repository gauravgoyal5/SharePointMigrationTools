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
    //using OfficeDevPnP.Core.Framework.Provisioning.Model;
    using OfficeDevPnP.Core.Pages;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security;
    using System.Text;
    using System.Threading.Tasks;
    using MigrationFactory.O365Groups.Model;
    using Microsoft.SharePoint.Client.WebParts;
    using OfficeDevPnP.Core.Entities;
    using Microsoft.SharePoint.Client.Utilities;
    public class SecurityPrincipalOperations
    {
        //public static string FindGroup(Web targetWeb, string loginName)
        //{
        //    var groupName = DoesGroupExist(targetWeb, loginName);

        //    if (groupName != string.Empty)
        //        return groupName;
        //    else
        //    {
        //        var loginNameWords = loginName.Split(' ');
        //        var newGroupName = targetWeb.SiteGroups.FirstOrDefault(g => g.LoginName.Contains(loginNameWords.Last())).LoginName;

        //        return DoesGroupExist(targetWeb, newGroupName);
        //    }
        //}

        //private static string DoesGroupExist(Web targetWeb, string loginName)
        //{
        //    var doesGroupExists = targetWeb.GroupExists(loginName);
        //    if (doesGroupExists)
        //        return loginName;
        //    else
        //        return string.Empty;
        //}
    }
}
