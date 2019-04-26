// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.Logging
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
namespace MigrationFactory.O365Groups.Logging
{
    using System;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface IAsyncLogger
    /// </summary> 
    public interface IAsyncLogger
    {
        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <returns>Task.</returns>
        Task Log(string message);
        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="args">The arguments.</param>
        /// <returns>Task.</returns>
        Task Log(string message, params object[] args);
        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="ex">The ex.</param>
        /// <returns>Task.</returns>
        Task Log(string message, Exception ex);

        /// <summary>
        /// Logs the warning.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <returns>Task.</returns>
        Task LogWarning(string message);
        /// <summary>
        /// Logs the warning.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="args">The arguments.</param>
        /// <returns>Task.</returns>
        Task LogWarning(string message, params object[] args);

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <returns>Task.</returns>
        Task LogError(string message);
        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="args">The arguments.</param>
        /// <returns>Task.</returns>
        Task LogError(string message, params object[] args);
    }
}
