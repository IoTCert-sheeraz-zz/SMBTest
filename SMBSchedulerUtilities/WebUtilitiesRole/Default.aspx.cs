using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using EWSAPIWrapper;
using System.Collections.Specialized;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Xml.Linq;
using System.IO;
using System.Web.Hosting;


namespace WebUtilitiesRole
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string xmlFilePath = Path.Combine(HostingEnvironment.ApplicationPhysicalPath, @"App_Data\DLList.xml");
                XDocument xDoc = XDocument.Load(xmlFilePath);
                ddlTimeZones.DataSource = xDoc.Root.Elements("DL").Attributes("ServiceAgentTimeZone").Select(item => item.Value).Distinct();
                ddlTimeZones.DataBind();
            }
        }


        protected List<AgentDLInfo> GetListDetails(string toBeCheckedDL, TimeSpan toBeCheckedOffset, DateTime startDate, DateTime endDate)
        {
            List<AgentDLInfo> agentDLList = new List<AgentDLInfo>();
            try
            {
                EWSUtility ewsUtility = new EWSUtility("SMBSPWEB", "]#Uun6~QqN7k@Z}c", "REDMOND", "smbspweb@microsoft.com");
                List<string> lstEmailID = (List<string>)ewsUtility.ExpandDistributionLists(toBeCheckedDL);
               // List<string> lstEmailID = new List<string>();
               // lstEmailID.Clear();
                //lstEmailID.Add("v-reorja@microsoft.com");
                //lstEmailID.Add("v-keodio@microsoft.com");
                //lstEmailID.Add("v-megari@microsoft.com");
                //lstEmailID.Add("v - anmoga@microsoft.com");
                //lstEmailID.Add("v - growen@microsoft.com");
                //lstEmailID.Add("v - mibude@microsoft.com");
                //lstEmailID.Add("v - rashaa@microsoft.com");

                //lstEmailID.Add("v-maross@microsoft.com");
                //lstEmailID.Add("v-50chsc@microsoft.com");


                // Temporary fix : The below agent is causing some issue while fetching data from Exchange, The error is "The XML document ended unexpectedly."
                // The affected team are : MastOpt and CoachNA (since both DLs have the same list of agents)
                // lstEmailID.Remove("v-10cabr@microsoft.com"); //todo
                //--End temporary fix 16 Feb 2017

                foreach (string agentEmailID in lstEmailID)
                {
                    GetUserAvailabilityResults freeBusyResults = ewsUtility.Exchange.GetUserAvailability(
                        Enumerable.Repeat(new AttendeeInfo { SmtpAddress = agentEmailID, AttendeeType = MeetingAttendeeType.Required }, 1),
                        new TimeWindow(startDate, endDate), AvailabilityData.FreeBusy,
                        new AvailabilityOptions() { MeetingDuration = 30, RequestedFreeBusyView = FreeBusyViewType.FreeBusy });

                    AgentDLInfo agentDLInfo = new AgentDLInfo();
                    agentDLInfo.DLName = toBeCheckedDL;
                    agentDLInfo.EmailId = agentEmailID;

                    if (freeBusyResults.AttendeesAvailability.First().WorkingHours != null)
                    {
                        agentDLInfo.TimeZoneInfo = freeBusyResults.AttendeesAvailability.First().WorkingHours.TimeZone;
                        agentDLInfo.AgentOffset = agentDLInfo.TimeZoneInfo.BaseUtcOffset.ToString();
                        agentDLInfo.DLOffset = toBeCheckedOffset.ToString();
                        agentDLInfo.IsBaseOffset = agentDLInfo.TimeZoneInfo.BaseUtcOffset.Equals(toBeCheckedOffset);
                        agentDLInfo.WorkhoursStartTime = freeBusyResults.AttendeesAvailability.First().WorkingHours.StartTime.ToString();
                        agentDLInfo.WorkhoursEndTime = freeBusyResults.AttendeesAvailability.First().WorkingHours.EndTime.ToString();
                    }
                    else
                    {
                        agentDLInfo.DLOffset = toBeCheckedOffset.ToString();
                        agentDLInfo.AgentOffset = "Unknown";
                        agentDLInfo.TimeZoneInfo = null;
                        agentDLInfo.IsBaseOffset = false;
                    }

                    agentDLList.Add(agentDLInfo);
                }

                return agentDLList;
            }
            catch (Exception ex)
            {
                lblError.Text = "Some Exception occurs due to invalid DL or due to some another issue.";
                lblError.Visible = true;
                return agentDLList;
            }

        }

        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            lblError.Visible = false;
            string standardName = ddlTimeZones.SelectedItem.Text;
            TimeSpan offset = TimeZoneInfo.FindSystemTimeZoneById(standardName).BaseUtcOffset;

            List<AgentDLInfo> agentDLList = GetListDetails(ddlDLNames.SelectedValue, offset, DateTime.UtcNow.AddDays(1), DateTime.UtcNow.AddDays(10));
            
            grdDLDetails.DataSource = agentDLList;
            grdDLDetails.DataBind();
        }
    }
}