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
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    /// <summary>
    /// Class MaximumRetryAttemptedException.
    /// Implements the <see cref="System.Exception" />
    /// </summary>
    /// <seealso cref="System.Exception" />
    [Serializable]
    public class MaximumRetryAttemptedException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MaximumRetryAttemptedException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public MaximumRetryAttemptedException(string message) : base(message) { }
    }
}
