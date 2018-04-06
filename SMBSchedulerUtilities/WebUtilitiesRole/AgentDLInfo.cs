using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebUtilitiesRole
{
    public class AgentDLInfo
    {
        public string DLName { get; set; }
        public TimeZoneInfo TimeZoneInfo { get; set; }
        public string AgentOffset { get; set; }

        public string DLOffset { get; set; }
        public string EmailId { get; set; }
        public bool IsBaseOffset { get; set; }

        public string WorkhoursStartTime { get; set; }
        public string WorkhoursEndTime { get; set; }
        public AgentDLInfo()
        { }
    }
}