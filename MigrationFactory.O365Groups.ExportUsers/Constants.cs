// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.ExportUsers
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
namespace MigrationFactory.O365Groups.ExportUsers
{
    /// <summary>
    /// Class Constants.
    /// </summary>
    public static class Constants
    {        
        /// <summary>
        /// The password message source
        /// </summary>
        public const string PasswordMessageSource = "Please provide password for Source Sites";
        
        /// <summary>
        /// Class AppSettings.
        /// </summary>
        public static class AppSettings
        {
            /// <summary>
            /// The site map key
            /// </summary>
            public const string UserExportSiteMapKey = "UserExportSiteMap";
            /// <summary>
            /// The site map sheet key
            /// </summary>
            public const string SiteMapSheetKey = "SiteMapSheetName";
            
            /// <summary>
            /// The logger instance key
            /// </summary>
            public const string LoggerInstanceKey = "ModernPageMigrationFileLogger";
            
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
            /// <summary>
            /// The retry count key
            /// </summary>
            public const string UserExportSiteMapReportKey = "ExportReportFile";
            /// <summary>
            /// The retry count key
            /// </summary>
            public const string DomainToSearchKey = "DomainToSearch";
            /// <summary>
            /// The batch size key
            /// </summary>
            public const string BatchSizeKey = "BatchSize";
        }

    }
}
