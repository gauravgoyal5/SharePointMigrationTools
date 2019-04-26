// ***********************************************************************
// Assembly         : MigrationFactory.O365Groups.Factory
// Author           : shiv
// Created          : 12-21-2018
//
// Last Modified By : shiv
// Last Modified On : 01-04-2019
// ***********************************************************************
namespace MigrationFactory.O365Groups.Factory
{
    using Logging;
    using System;
    using System.Configuration;


    /// <summary>
    /// Class O365GroupsFactory.
    /// </summary>
    public class O365GroupsFactory
    {
        /// <summary>
        /// Creates the specified object type.
        /// </summary>
        /// <param name="objType">Type of the object.</param>
        /// <returns>Object.</returns>
        public static Object Create(ObjectType objType)
        {
            switch (objType)
            {
                case ObjectType.CSVOperation:
                    

                case ObjectType.WebOperation:
                    

                default:
                    return null;
                    
            }
        }


        /// <summary>
        /// Gets the logger.
        /// </summary>
        /// <param name="loggerSource">The logger source.</param>
        /// <returns>Object.</returns>
        public static Object GetLogger(string[] loggerSource)
        {
            IAsyncLogger logger = new TraceSourceLogger(loggerSource);

            return logger;
        }
    }

    /// <summary>
    /// Enum ObjectType
    /// </summary>
    public enum ObjectType
    {
        /// <summary>
        /// The external user repository type
        /// </summary>
        WebOperation,
        /// <summary>
        /// The logger tyoe
        /// </summary>
        Logger,

        /// <summary>
        /// The cache repository type
        /// </summary>
        CSVOperation
    }

}
