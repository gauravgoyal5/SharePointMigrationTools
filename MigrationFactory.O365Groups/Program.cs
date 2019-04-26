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
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using MigrationFactory.O365Groups.Model;
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
            Console.WriteLine("Please enter the user name.");
            string userName = Console.ReadLine(); //ConfigurationManager.AppSettings["UserName"];

            Console.WriteLine("Please enter the password.");
            string txtPassword = Console.ReadLine(); //ConfigurationManager.AppSettings["Password"];

            string webUrl = ConfigurationManager.AppSettings["WebUrl"];
            string groupExportPath = ConfigurationManager.AppSettings["GroupExportPath"];
            string groupExportSheetName = ConfigurationManager.AppSettings["GroupExportSheetName"];

            var csvOperation = new CSVOperations();
            var groupDetails = csvOperation.ReadFile("GroupExport", groupExportPath, groupExportSheetName).Cast<GroupExportReport>().ToList();
            
            if(groupDetails != null)
            {
                foreach (var group in groupDetails)
                {
                    if (!string.IsNullOrEmpty(group.Id))
                    {
                        Console.WriteLine("Processing " + group.DisplayName);
                        var Id = group.Id.Substring(group.DisplayName.Length + 1);
                        ClaimsWebClient wc = new ClaimsWebClient(new Uri(webUrl), userName, txtPassword);
                        var fileRelativeUrl = Constants.GROUP_STATUS_URL + $"?id={Id}&target=site"; //https://tenantname.sharepoint.com/_layouts/groupstatus.aspx?id=567da1b0-8a75-4405-9553-12e8c30c1234&target=site 

                        byte[] response = wc.DownloadData(webUrl + fileRelativeUrl);
                        Console.WriteLine("Processed " + group.DisplayName);
                    }
                }
            }

            Console.WriteLine("Operation Completed Successfully!");

            Console.Read();
            
        }
    }
}
