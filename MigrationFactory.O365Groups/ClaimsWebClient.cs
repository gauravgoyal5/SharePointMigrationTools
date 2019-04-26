// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************

namespace MigrationFactory.O365Groups
{
    using Microsoft.SharePoint.Client;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Security;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Class ClaimsWebClient.
    /// Implements the <see cref="System.Net.WebClient" />
    /// </summary>
    /// <seealso cref="System.Net.WebClient" />
    public class ClaimsWebClient : WebClient
    {
        /// <summary>
        /// The cookie container
        /// </summary>
        private CookieContainer CookieContainer;

        /// <summary>
        /// Initializes a new instance of the <see cref="ClaimsWebClient"/> class.
        /// </summary>
        /// <param name="host">The host.</param>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">The password.</param>
        public ClaimsWebClient(Uri host, string userName, string password)
        {
            CookieContainer = GetAuthCookies(host, userName, password);
        }
        /// <summary>
        /// Returns a <see cref="T:System.Net.WebRequest" /> object for the specified resource.
        /// </summary>
        /// <param name="address">A <see cref="T:System.Uri" /> that identifies the resource to request.</param>
        /// <returns>A new <see cref="T:System.Net.WebRequest" /> object for the specified resource.</returns>
        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest request = base.GetWebRequest(address);
            if (request is HttpWebRequest)
            {
                (request as HttpWebRequest).CookieContainer = CookieContainer;
            }

            return request;
        }
        /// <summary>
        /// Gets the authentication cookies.
        /// </summary>
        /// <param name="webUri">The web URI.</param>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">The password.</param>
        /// <returns>CookieContainer.</returns>
        private static CookieContainer GetAuthCookies(Uri webUri, string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (var c in password) { securePassword.AppendChar(c); }
            var credentials = new SharePointOnlineCredentials(userName, securePassword);
            var authCookie = credentials.GetAuthenticationCookie(webUri);
            var cookieContainer = new CookieContainer();
            cookieContainer.SetCookies(webUri, authCookie); return cookieContainer;
        }

    }
}
