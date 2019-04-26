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
    using Microsoft.SharePoint.Client;
    using System;
    using System.Net;
    using System.Threading;
    /// <summary>
    /// Class ThrottleHandler.
    /// </summary>
    public static class ThrottleHandler
    {
        /// <summary>
        /// Executes the query with incremental retry.
        /// </summary>
        /// <param name="clientContext">The client context.</param>
        /// <param name="retryCount">The retry count.</param>
        /// <param name="delay">The delay.</param>
        /// <exception cref="ArgumentException">
        /// Provide a retry count greater than zero.
        /// or
        /// Provide a delay greater than zero.
        /// </exception>
        /// <exception cref="MaximumRetryAttemptedException">Maximum retry attempts {retryCount}</exception>
        public static void ExecuteQueryWithIncrementalRetry(this ClientContext clientContext, int retryCount, int delay)
        {
            int retryAttempts = 0;
            int backoffInterval = delay;
            int retryAfterInterval = 0;
            bool retry = false;
            ClientRequestWrapper wrapper = null;
            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");
            if (delay <= 0)
                throw new ArgumentException("Provide a delay greater than zero.");

            clientContext.RequestTimeout = Timeout.Infinite;

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
                {
                    if (!retry)
                    {
                        clientContext.ExecuteQuery();
                        return;
                    }
                    else
                    {
                        // retry the previous request
                        if (wrapper != null && wrapper.Value != null)
                        {
                            clientContext.RetryQuery(wrapper.Value);
                            return;
                        }
                    }
                }
                catch (WebException ex)
                {
                    var response = ex.Response as HttpWebResponse;
                    // Check if request was throttled - http status code 429
                    // Check is request failed due to server unavailable - http status code 503
                    if (response != null && (response.StatusCode == (HttpStatusCode)429 || response.StatusCode == (HttpStatusCode)503))
                    {
                        wrapper = (ClientRequestWrapper)ex.Data["ClientRequest"];
                        retry = true;

                        // Determine the retry after value - use the retry-after header when available
                        string retryAfterHeader = response.GetResponseHeader("Retry-After");
                        if (!string.IsNullOrEmpty(retryAfterHeader))
                        {
                            if (!Int32.TryParse(retryAfterHeader, out retryAfterInterval))
                            {
                                retryAfterInterval = backoffInterval;
                            }
                        }
                        else
                        {
                            retryAfterInterval = backoffInterval;
                        }

                        // Delay for the requested seconds
                        Thread.Sleep(retryAfterInterval * 1000);

                        // Increase counters
                        retryAttempts++;
                        backoffInterval = backoffInterval * 2;
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            throw new MaximumRetryAttemptedException($"Maximum retry attempts {retryCount}, has be attempted.");
        }

        
    }
}
