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
    using System.Diagnostics;

    /// <summary>
    /// Class TraceSourceLogger.
    /// </summary>
    /// <seealso cref="MigrationFactory.O365Groups.Logging.IAsyncLogger" />
    public class TraceSourceLogger : IAsyncLogger
    {
        /// <summary>
        /// The trace source
        /// </summary>
        private readonly TraceSource[] traceSourceArr;

        /// <summary>
        /// Initializes a new instance of the <see cref="TraceSourceLogger" /> class.
        /// </summary>
        /// <param name="source">The source.</param>
        public TraceSourceLogger(string[] source)
        {
            this.traceSourceArr = new TraceSource[source.Length];
            for (int i= 0; i< source.Length; i++)
            {
                this.traceSourceArr.SetValue(new TraceSource(source[i]), i);
            }
        }

        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <returns>Task.</returns>
        public async Task Log(string message)
        {
            await Task.Run(() =>
            {
                foreach (TraceSource traceSource in traceSourceArr)
                {
                    traceSource.TraceInformation(message);
                }
            });
        }

        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="args">The arguments.</param>
        /// <returns>Task.</returns>
        public async Task Log(string message, params object[] args)
        {
            await Task.Run(() =>
            {
                foreach (TraceSource traceSource in traceSourceArr)
                {
                    traceSource.TraceInformation(message, args);
                }
            });
        }

        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="ex">The ex.</param>
        /// <returns>Task.</returns>
        public async Task Log(string message, Exception ex)
        {
            await Task.Run(() =>
            {
                foreach (TraceSource traceSource in traceSourceArr)
                {
                    traceSource.TraceData(TraceEventType.Error, 9801, $"{message}" + Environment.NewLine + ex);
                }
            });
        }

        /// <summary>
        /// Logs the warning.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <returns>Task.</returns>
        public async Task LogWarning(string message)
        {
            await Task.Run(() =>
            {
                foreach (TraceSource traceSource in traceSourceArr)
                {
                    traceSource.TraceEvent(TraceEventType.Warning, 9800, message);
                }
            });
        }

        /// <summary>
        /// Logs the warning.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="args">The arguments.</param>
        /// <returns>Task.</returns>
        public async Task LogWarning(string message, params object[] args)
        {
            await Task.Run(() =>
            {
                foreach (TraceSource traceSource in traceSourceArr)
                {
                    traceSource.TraceEvent(TraceEventType.Warning, 9800, message, args);
                }
            });
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <returns>Task.</returns>
        public async Task LogError(string message)
        {
            await Task.Run(() =>
            {
                foreach (TraceSource traceSource in traceSourceArr)
                {
                    traceSource.TraceEvent(TraceEventType.Error, 9801, message);
                }
            });
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="args">The arguments.</param>
        /// <returns>Task.</returns>
        public async Task LogError(string message, params object[] args)
        {
            await Task.Run(() =>
            {
                foreach (TraceSource traceSource in traceSourceArr)
                {
                    traceSource.TraceEvent(TraceEventType.Error, 9801, message, args);
                }
            });
        }
    }
}