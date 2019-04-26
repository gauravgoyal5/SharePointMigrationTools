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
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface IFactory
    /// </summary>
    public interface IFactory
    {
        /// <summary>
        /// Creates the specified object type.
        /// </summary>
        /// <param name="objType">Type of the object.</param>
        /// <returns>Object.</returns>
        Object Create(ObjectType objType);

        /// <summary>
        /// Gets the logger.
        /// </summary>
        /// <param name="objType">Type of the object.</param>
        /// <param name="loggerSource">The logger source.</param>
        /// <returns>Object.</returns>
        Object GetLogger(ObjectType objType, string loggerSource);
    }
}
