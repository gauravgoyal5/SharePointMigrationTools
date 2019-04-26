// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.ModernPage
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
namespace MigrationFactory.O365Groups.ModernPage.Entity
{
    using Microsoft.SharePoint.Client;
    using System.Security;

    /// <summary>
    /// Class ModernWeb.
    /// </summary>
    public class ModernWeb
    {
        /// <summary>
        /// Gets or sets the source site URL.
        /// </summary>
        /// <value>The source site URL.</value>
        public string SourceSiteUrl { get; set; }
        /// <summary>
        /// Gets or sets the target site URL.
        /// </summary>
        /// <value>The target site URL.</value>
        public string TargetSiteUrl { get; set; }
        /// <summary>
        /// Gets or sets the name of the source user.
        /// </summary>
        /// <value>The name of the source user.</value>
        public string SourceUserName { get; set; }
        /// <summary>
        /// Gets or sets the name of the target user.
        /// </summary>
        /// <value>The name of the target user.</value>
        public string TargetUserName { get; set; }
        /// <summary>
        /// Gets or sets the source password.
        /// </summary>
        /// <value>The source password.</value>
        public SecureString SourcePassword { get; set; }
        /// <summary>
        /// Gets or sets the target password.
        /// </summary>
        /// <value>The target password.</value>
        public SecureString TargetPassword { get; set; }
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


    }
}
