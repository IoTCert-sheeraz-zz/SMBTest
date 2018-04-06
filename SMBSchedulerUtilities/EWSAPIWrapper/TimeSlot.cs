// --------------------------------------------------------------------------
// <copyright file="TimeSlot.cs" company="Microsoft Corporation">
//   Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace EWSAPIWrapper
{
    using System;

    /// <summary>
    /// This Class represent a free/busy time slot of Service Agent Working Hours.
    /// </summary>
  public class TimeSlot
    {
        /// <summary>
        /// Gets or sets the start time of slot with in working hours 
        /// of the Service Agent Working Hours.
        /// </summary>
        /// <value>
        /// The start time.
        /// </value>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// Gets or sets the end time of slot with in working hours
        /// of the Service Agent Working Hours.
        /// </summary>
        /// <value>
        /// The end time.
        /// </value>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [free slot].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [free slot]; otherwise, <c>false</c>.
        /// </value>
        public bool FreeSlot { get; set; }
    }
}
