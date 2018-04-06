// --------------------------------------------------------------------------
// <copyright file="AppointmentInfo.cs" company="Microsoft Corporation">
//   Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------
namespace EWSAPIWrapper
{
    using Microsoft.Exchange.WebServices.Data;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This Class holds the information required to send a meeting appointment
    /// </summary>
    public class AppointmentInfo
    {

        /// <summary>
        /// Gets or sets the ID.
        /// </summary>
        /// <value>
        /// The ID.
        /// </value>
        public string ID { get; set; }

        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        /// <value>
        /// The subject.
        /// </value>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets the start.
        /// </summary>
        /// <value>
        /// The start.
        /// </value>
        public DateTime Start { get; set; }

        /// <summary>
        /// Gets or sets the end.
        /// </summary>
        /// <value>
        /// The end.
        /// </value>
        public DateTime End { get; set; }

        /// <summary>
        /// Gets or sets the Update Meeting Time.
        /// </summary>
        /// <value>
        /// The Update.
        /// </value>
        /// 
        public DateTime UpdateDate { get; set; }

        /// <summary>
        /// Gets or sets the Cancel Meeting Time.
        /// </summary>
        /// <value>
        /// The Cancel.
        /// </value>
        public DateTime CancelDate { get; set; }       

        /// <summary>
        /// Gets or sets the user ID.
        /// </summary>
        /// <value>
        /// The user ID.
        /// </value>
        public int? UserID { get; set; }

        /// <summary>
        /// Gets or sets the reminder.
        /// </summary>
        /// <value>
        /// The reminder.
        /// </value>
        public int Reminder { get; set; }

        /// <summary>
        /// Gets or sets the description.
        /// </summary>
        /// <value>
        /// The description.
        /// </value>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the customer mail id.
        /// </summary>
        /// <value>
        /// The customer mail id.
        /// </value>
        public string CustomerMailId { get; set; }

        /// <summary>
        /// Gets or sets the agent mail id.
        /// </summary>
        /// <value>
        /// The agent mail id.
        /// </value>
        public string AgentMailId { get; set; }        

    }
}
