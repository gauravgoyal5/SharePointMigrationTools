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
    /// <summary>
    /// Class ConsoleOperations.
    /// </summary>
    class ConsoleOperations
    {
        /// <summary>
        /// Writes to console.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="color">The color.</param>
        public static void WriteToConsole(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
