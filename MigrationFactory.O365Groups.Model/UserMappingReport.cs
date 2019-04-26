// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.Model
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrationFactory.O365Groups.Model
{
    /// <summary>
    /// Class UserMappingReport.
    /// Implements the <see cref="MigrationFactory.O365Groups.Model.IReport" />
    /// </summary>
    /// <seealso cref="MigrationFactory.O365Groups.Model.IReport" />
    public class UserMappingReport: IReport
    {
        /// <summary>
        /// Gets or sets the identifier.
        /// </summary>
        /// <value>The identifier.</value>
        public string Id { get; set; }
        /// <summary>
        /// Gets or sets the source user identifier.
        /// </summary>
        /// <value>The source user identifier.</value>
        public string SourceUserId { get; set; }
        /// <summary>
        /// Gets or sets the target user identifier.
        /// </summary>
        /// <value>The target user identifier.</value>
        public string TargetUserId { get; set; }
    }
}
