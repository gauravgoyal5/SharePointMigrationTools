// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.Model
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
using DocumentFormat.OpenXml.Packaging;
namespace MigrationFactory.O365Groups.Model
{
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Threading.Tasks;
    /// <summary>
    /// Class CSVOperations.
    /// </summary>
    public class CSVOperations
    {
        /// <summary>
        /// Reads the file.
        /// </summary>
        /// <param name="reportType">Type of the report.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>List&lt;IReport&gt;.</returns>
        /// <exception cref="Exception">Could not find sheet with name {sheetName}</exception>
        public List<IReport> ReadFile(string reportType, string fileName, string sheetName)
        {
            string path = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, "Resource");

            using (var spreadsheetDocument = SpreadsheetDocument.Open(path + "\\" + fileName, false))
            {

                var workbookPart = spreadsheetDocument.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

                if (sheet == null)
                    throw new Exception($"Could not find sheet with name {sheetName}");

                var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                if (worksheetPart != null)
                {
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    List<IReport> reportList = null;

                    switch (reportType)
                    {
                        case "GroupExport":
                            var workflowReportsList = sheetData.Elements<Row>().Select(r => new GroupExportReport()
                            {
                                Id = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(0),
                                    sharedStringTablePart),
                                DisplayName = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(1),
                                    sharedStringTablePart),
                            }).ToList();

                            reportList = workflowReportsList.Cast<IReport>().ToList();
                            break;

                        case "SiteMap":
                            var siteMapList = sheetData.Elements<Row>().Select(r => new SiteMapReport()
                            {
                                Id = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(0),
                                    sharedStringTablePart),
                                SourceSiteUrl = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(1),
                                    sharedStringTablePart),
                                TargetSiteUrl = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(2),
                                    sharedStringTablePart),
                                SourceUser = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(3),
                                    sharedStringTablePart),
                                TargetUser = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(4),
                                    sharedStringTablePart)
                            }).ToList();

                            reportList = siteMapList.Cast<IReport>().ToList();
                            break;

                        case "UserExportSiteMap":
                            var userExportSiteMapList = sheetData.Elements<Row>().Select(r => new UserExportSiteMapReport()
                            {
                                Id = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(0),
                                    sharedStringTablePart),
                                SourceSiteUrl = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(1),
                                    sharedStringTablePart),
                                SourceUser = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(2),
                                    sharedStringTablePart)
                            }).ToList();

                            reportList = userExportSiteMapList.Cast<IReport>().ToList();
                            break;

                        case "UserMapping":
                            var userMappingList = sheetData.Elements<Row>().Select(r => new UserMappingReport()
                            {
                                Id = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(0),
                                    sharedStringTablePart),
                                SourceUserId = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(1),
                                    sharedStringTablePart),
                                TargetUserId = GetFormattedCellValue(r.Elements<Cell>().ElementAtOrDefault(2),
                                    sharedStringTablePart)
                            }).ToList();

                            reportList = userMappingList.Cast<IReport>().ToList();
                            break;

                        default:
                            break;
                    }

                    return reportList;
                }
                else
                {
                    return null;
                }

            }
        }

        /// <summary>
        /// Gets the formatted cell value.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="sharedStringTablePart">The shared string table part.</param>
        /// <returns>System.String.</returns>
        private string GetFormattedCellValue(Cell cell, SharedStringTablePart sharedStringTablePart)
        {
            var text = string.Empty;

            if (cell != null && cell.CellValue != null)
            {
                text = cell.CellValue.Text;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    text = sharedStringTablePart.SharedStringTable.ElementAt(int.Parse(text, CultureInfo.InvariantCulture)).InnerText;
                }
            }

            return text.Trim().ToLower();
        }

        /// <summary>
        /// Writes the CSV.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="items">The items.</param>
        /// <param name="path">The path.</param>
        public bool WriteCsv<T>(IEnumerable<T> items, string fileName)
        {
            bool isSuccess = false;
            try
            {                
                string path = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, "Resource") + "\\" + fileName;
                bool doesFileExists = File.Exists(path);
                Type itemType = typeof(T);
                var props = itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .OrderBy(p => p.Name);

                using (var writer = new StreamWriter(path, true))
                {
                    if(!doesFileExists)
                        writer.WriteLine(string.Join(", ", props.Select(p => p.Name)));

                    foreach (var item in items)
                    {
                        writer.WriteLine(string.Join(", ", props.Select(p => p.GetValue(item, null))));
                    }
                    isSuccess = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in Exporting users {ex.Message}");
            }

            return isSuccess;
        }
    }
}
