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
    /// <summary>
    /// Class Constants.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// The web template
        /// </summary>
        public const string WebTemplate = "GROUP";
        /// <summary>
        /// The modern page library
        /// </summary>
        public const string ModernPageLibrary = "Site Pages";

        /// <summary>
        /// The modern page content type
        /// </summary>
        public const string ModernPageContentType = "Site Page";
        /// <summary>
        /// The modern page content control
        /// </summary>
        public const string ModernPageContentControl = "CanvasContent1";

        /// <summary>
        /// The wiki page content type
        /// </summary>
        public const string WikiPageContentType = "Wiki Page";
        /// <summary>
        /// The wiki page content control
        /// </summary>
        public const string WikiPageContentControl = "WikiField";

        /// <summary>
        /// The repost page content type
        /// </summary>
        public const string RepostPageContentType = "Repost Page";
        /// <summary>
        /// The web part page content type
        /// </summary>
        public const string WebPartPageContentType = "Web Part Page";

        /// <summary>
        /// The file reference
        /// </summary>
        public const string FileRef = "FileRef";

        /// <summary>
        /// The PageLayoutType reference
        /// </summary>
        public const string PageLayoutType = "PageLayoutType";

        /// <summary>
        /// The ClientSideApplicationId reference
        /// </summary>
        public const string ClientSideApplicationId = "ClientSideApplicationId";

        /// <summary>
        /// The LayoutWebpartsContent reference
        /// </summary>
        public const string LayoutWebpartsContent = "LayoutWebpartsContent";

        /// <summary>
        /// The first published date
        /// </summary>
        public const string FirstPublishedDate = "FirstPublishedDate";
        /// <summary>
        /// The banner image URL
        /// </summary>
        public const string BannerImageUrl = "BannerImageUrl";
        /// <summary>
        /// The content type identifier
        /// </summary>
        public const string ContentTypeId = "ContentTypeId";
        /// <summary>
        /// The description
        /// </summary>
        public const string Description = "Description";
        /// <summary>
        /// The original source URL
        /// </summary>
        public const string OriginalSourceUrl = "_OriginalSourceUrl";
        /// <summary>
        /// The promoted state
        /// </summary>
        public const string PromotedState = "PromotedState";
        /// <summary>
        /// The author
        /// </summary>
        public const string Author = "Author";
        /// <summary>
        /// The editor
        /// </summary>
        public const string Editor = "Editor";
        /// <summary>
        /// The created
        /// </summary>
        public const string Created = "Created";
        /// <summary>
        /// The modified
        /// </summary>
        public const string Modified = "Modified";
        /// <summary>
        /// The file leaf reference
        /// </summary>
        public const string FileLeafRef = "FileLeafRef";
        /// <summary>
        /// The site map sheet
        /// </summary>
        public const string SiteMapSheet = "SiteMap";
        /// <summary>
        /// The user mapping sheet
        /// </summary>
        public const string UserMappingSheet = "UserMapping";
        /// <summary>
        /// The password message source
        /// </summary>
        public const string PasswordMessageSource = "Please provide password for Source Sites";
        /// <summary>
        /// The password message target
        /// </summary>
        public const string PasswordMessageTarget = "Please provide password for Target Sites";
        /// <summary>
        /// CAML query to extract List of Modern Pages
        /// </summary>
        public const string ModernPageQueryString = "<View>"
                                                   + "<Query>"
                                                   + "<Where><Eq><FieldRef Name='ContentType' /><Value Type='Computed'>Site Page</Value></Eq></Where>"
                                                   + "</Query>"
                                                   + "</View>";

        /// <summary>
        /// Class AppSettings.
        /// </summary>
        public static class AppSettings
        {
            /// <summary>
            /// The site map key
            /// </summary>
            public const string SiteMapKey = "SiteMap";
            /// <summary>
            /// The site map sheet key
            /// </summary>
            public const string SiteMapSheetKey = "SiteMapSheetName";
            /// <summary>
            /// The user mapping key
            /// </summary>
            public const string UserMappingKey = "UserMapping";
            /// <summary>
            /// The user mapping sheet key
            /// </summary>
            public const string UserMappingSheetKey = "UserMappingSheetName";
            /// <summary>
            /// The logger instance key
            /// </summary>
            public const string LoggerInstanceKey = "ModernPageMigrationFileLogger";
            /// <summary>
            /// The migration type key
            /// </summary>
            public const string MigrationTypeKey = "MigrationType";
            /// <summary>
            /// The default user key
            /// </summary>
            public const string DefaultUserKey = "DefaultUser";
            /// <summary>
            /// The user agent key
            /// </summary>
            public const string UserAgentKey = "UserAgent";
            /// <summary>
            /// The delay key
            /// </summary>
            public const string DelayKey = "Delay";
            /// <summary>
            /// The retry count key
            /// </summary>
            public const string RetryCountKey = "RetryCount";
        }

    }
}
