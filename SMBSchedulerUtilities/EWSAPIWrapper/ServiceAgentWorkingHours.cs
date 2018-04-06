// --------------------------------------------------------------------------
// <copyright file="ServiceAgentWorkingHours.cs" company="Microsoft Corporation">
//   Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace EWSAPIWrapper
{
    using System;

    /// <summary>
    /// This class represents the working hours of a Service Agent & 
    /// time 
    /// </summary>
    public class ServiceAgentWorkingHours
    {
        /// <summary>
        /// Gets or sets the start time of the Service Agent Shift.
        /// </summary>
        /// <value>
        /// The start time.
        /// </value>        
        public TimeSpan StartTime { get; set; }

        /// <summary>
        /// Gets or sets the end time of the Service Agent Shift.
        /// </summary>
        /// <value>
        /// The end time.
        /// </value>        
        public TimeSpan EndTime { get; set; }

        /// <summary>
        /// Gets or sets the service agent time zone.
        /// </summary>
        /// <value>
        /// The service agent time zone.
        /// </value>        
        public string ServiceAgentTimeZone { get; set; }
    }
}
