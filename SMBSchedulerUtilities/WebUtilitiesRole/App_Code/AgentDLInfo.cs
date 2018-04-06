using System;

namespace WebUtilitiesRole
{
    public class AgentDLInfo
    {
        public string DLName { get; set; }
        public TimeZoneInfo TimeZoneInfo { get; set; }
        public string Offset { get; set; }
        public string EmailId { get; set; }
        public bool IsBaseOffset { get; set; }

        public string WorkhoursStartTime { get; set; }
        public string WorkhoursEndTime { get; set; }
        public AgentDLInfo()
        { }
    }
}