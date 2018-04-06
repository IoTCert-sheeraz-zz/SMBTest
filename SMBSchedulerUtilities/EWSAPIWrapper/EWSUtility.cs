// --------------------------------------------------------------------------
// <copyright file="EWSUtility.cs" company="Microsoft Corporation">
//   Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace EWSAPIWrapper
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using Microsoft.Exchange.WebServices.Data;
    using System.Collections.ObjectModel;

    /// <summary>
    /// This class provides facility to get Microsoft Exchange data using EWS API 2.0
    /// </summary>
  public class EWSUtility
    {
        /// <summary>
        /// Gets or sets the exchange.
        /// </summary>
        /// <value>
        /// The exchange.
        /// </value>
      public ExchangeService Exchange { get; set; }

      /// <summary>
      /// Gets or sets the requested time window.
      /// </summary>
      /// <value>
      /// The requested time window.
      /// </value>
      public TimeWindow RequestedTimeWindow { get; set; }

      /// <summary>
      /// Gets or sets the distribution list.
      /// </summary>
      /// <value>
      /// The distribution list.
      /// </value>
      public string DistributionList { get; set; }

      /// <summary>
      /// Initializes a new instance of the <see cref="EWSUtility"/> class.
      /// </summary>
      /// <param name="userName">Name of the user.</param>
      /// <param name="password">The password.</param>
      /// <param name="domain">The domain.</param>
      /// <param name="email">The email.</param>
      public EWSUtility(string userName, string password, string domain,string email)
      {
          try
          {
              ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Utc);              
              service.Credentials = new NetworkCredential(email, password);              
              service.Url = new Uri("https://outlook.office365.com/ews/exchange.asmx");

              //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010, TimeZoneInfo.Utc);
              // service.Credentials = new NetworkCredential(userName, password, domain);
              //service.AutodiscoverUrl(email);

              Exchange = service;   
          }
          catch
          {
              throw;
          }
         
      }


      /// <summary>
      /// Expands the distribution lists.
      /// </summary>
      /// <param name="target">The target.</param>
      /// <returns>IReadOnlyList</returns>
      public IReadOnlyList<string> ExpandDistributionLists(string target)
      {
          try
          {
              // Return the expanded group.
              this.DistributionList = target;
              ExpandGroupResults myGroupMembers = this.Exchange.ExpandGroup(this.DistributionList);
              var emailCollection = (from x in myGroupMembers.Members
                                     select x.Address);
              return emailCollection.ToList();
          }
          catch (Exception ex)
          {
              throw ex;
          }
      }


      public IReadOnlyList<ServiceAgent> GetUserFreeBusy(IReadOnlyList<string> agentIDs, string targetTimeZone, TimeSpan shiftStartTime, TimeSpan shiftEndTime, DateTime startDate, DateTime endDate, out int rehitcount ,out bool isExchangeFailure)
      {
          try
          {
              rehitcount = 0;
              isExchangeFailure = false;
              RequestedTimeWindow = new TimeWindow(startDate, endDate);
              List<string> missedAgentIDs = new List<string>();
              Collection<TimeSlot> lstTimeSlot;
              Collection<ServiceAgent> lstServiceAgent = new Collection<ServiceAgent>();
              // Create a list of attendees.
              Collection<AttendeeInfo> attendees = new Collection<AttendeeInfo>();
              Collection<AttendeeInfo> attendeesInternal = new Collection<AttendeeInfo>();
              bool exchangeDataRetrival = true;
              int exchangeRehit = 0;
              int count = 0;
              int totalcount = 0;
              int agentsQueried = 0;
              int agentsQueriedMissed = 0;
              int agentsRetrieved = 0;
              if (agentIDs != null || agentIDs.Count > 0)
              {
                  foreach (var agentID in agentIDs)
                  {
                      count += 1;
                      totalcount += 1;
                      attendees.Add(new AttendeeInfo()
                      {
                          SmtpAddress = agentID,
                          AttendeeType = MeetingAttendeeType.Required
                      });

                      if ((count == 13) || (totalcount == agentIDs.Count))
                      {
                          count = 0;
                          while (exchangeDataRetrival)
                          {
                              missedAgentIDs.Clear();
                              // Specify availability options.
                              AvailabilityOptions myOptions = new AvailabilityOptions();
                              myOptions.MeetingDuration = 30;
                              myOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;
                              agentsQueried = attendees.Count;


                              // Return a set of free/busy times.
                              GetUserAvailabilityResults freeBusyResults = this.Exchange.GetUserAvailability(attendees,
                                                                                                   this.RequestedTimeWindow,
                                                                                                       AvailabilityData.FreeBusy,
                                                                                                    myOptions);
                              int i = 0;
                              foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
                              {
                                  if (availability.WorkingHours != null)
                                  {
                                      ServiceAgent srAgent = new ServiceAgent();
                                      srAgent.ServiceAgentEmail = attendees[i++].SmtpAddress;
                                      ServiceAgentWorkingHours srAgntWorkingHours = new ServiceAgentWorkingHours();
                                      srAgntWorkingHours.StartTime = availability.WorkingHours.StartTime;
                                      srAgntWorkingHours.EndTime = availability.WorkingHours.EndTime;
                                      srAgntWorkingHours.ServiceAgentTimeZone = targetTimeZone;
                                      srAgent.ServiceWorkingHours = srAgntWorkingHours;
                                      lstTimeSlot = new Collection<TimeSlot>();
                                      foreach (CalendarEvent calendarItem in availability.CalendarEvents)
                                      {
                                          if (calendarItem.FreeBusyStatus.ToString() != "Free")
                                          {
                                              TimeSlot tmSlot = new TimeSlot();
                                              tmSlot.StartTime = TimeZoneInfo.ConvertTimeFromUtc(calendarItem.StartTime, TimeZoneInfo.FindSystemTimeZoneById(targetTimeZone));
                                              tmSlot.EndTime = TimeZoneInfo.ConvertTimeFromUtc(calendarItem.EndTime, TimeZoneInfo.FindSystemTimeZoneById(targetTimeZone));
                                              tmSlot.FreeSlot = false;
                                              lstTimeSlot.Add(tmSlot);
                                          }
                                      }
                                      srAgent.BusyTimeSlotsCustomerTimeZone = lstTimeSlot;
                                      lstServiceAgent.Add(srAgent);
                                  }
                                  else
                                  {
                                      missedAgentIDs.Add(attendees[i].SmtpAddress);
                                      i++;
                                  }
                              }

                              agentsQueriedMissed = missedAgentIDs.Count;

                              if (missedAgentIDs.Count > 0 && exchangeRehit == 0)
                              {
                                  attendees.Clear();
                                  foreach (var agentMissed in missedAgentIDs)
                                  {
                                      attendees.Add(new AttendeeInfo()
                                      {
                                          SmtpAddress = agentMissed,
                                          AttendeeType = MeetingAttendeeType.Required
                                      });
                                  }
                                  exchangeRehit += 1;
                                  rehitcount = exchangeRehit;
                                 // isExchangeFailure = false;
                              }
                              else if (missedAgentIDs.Count > 0 && exchangeRehit != 0)
                              {
                                  
                                  if((agentsQueried-agentsQueriedMissed == 0))
                                  {
                                      rehitcount = exchangeRehit;
                                      isExchangeFailure = true;
                                      break;
                                  }
                               
                                  attendees.Clear();
                                  foreach (var agentMissed in missedAgentIDs)
                                  {
                                      attendees.Add(new AttendeeInfo()
                                      {
                                          SmtpAddress = agentMissed,
                                          AttendeeType = MeetingAttendeeType.Required
                                      });
                                  }
                                  exchangeRehit += 1;

                              }
                              else if (missedAgentIDs.Count == 0)
                              {
                                  break;
                              }
                          }
                          attendees.Clear();
                      }

                  }
              }
              return lstServiceAgent;
          }
          catch
          {
              throw;
          }
      }




      //public IReadOnlyList<ServiceAgent> GetUserFreeBusy(IReadOnlyList<string> agentIDs, string targetTimeZone, TimeSpan shiftStartTime, TimeSpan shiftEndTime, DateTime startDate, DateTime endDate)
      //{
      //    try
      //    {
      //        RequestedTimeWindow = new TimeWindow(startDate, endDate);
      //        List<string> missedAgentIDs = new List<string>();
      //        Collection<TimeSlot> lstTimeSlot;
      //        Collection<ServiceAgent> lstServiceAgent = new Collection<ServiceAgent>();
      //        // Create a list of attendees.
      //        Collection<AttendeeInfo> attendees = new Collection<AttendeeInfo>();
      //        Collection<AttendeeInfo> attendeesInternal = new Collection<AttendeeInfo>();
      //        bool exchangeDataRetrival = true;
      //        int exchangeRehit = 0;
      //        int count = 0;
      //        int totalcount = 0;
      //        if (agentIDs != null || agentIDs.Count > 0)
      //        {
      //            foreach (var agentID in agentIDs)
      //            {
      //                count += 1;
      //                totalcount += 1;
      //                attendees.Add(new AttendeeInfo()
      //                {
      //                    SmtpAddress = agentID,
      //                    AttendeeType = MeetingAttendeeType.Required
      //                });

      //                if ((count == 13) || (totalcount == agentIDs.Count))
      //                {
      //                    count = 0;
      //                    while (exchangeDataRetrival)
      //                    {
      //                        missedAgentIDs.Clear();
      //                        // Specify availability options.
      //                        AvailabilityOptions myOptions = new AvailabilityOptions();
      //                        myOptions.MeetingDuration = 30;
      //                        myOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;



      //                        // Return a set of free/busy times.
      //                        GetUserAvailabilityResults freeBusyResults = this.Exchange.GetUserAvailability(attendees,
      //                                                                                             this.RequestedTimeWindow,
      //                                                                                                 AvailabilityData.FreeBusy,
      //                                                                                              myOptions);
      //                        int i = 0;
      //                        foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
      //                        {
      //                            if (availability.WorkingHours != null)
      //                            {
      //                                ServiceAgent srAgent = new ServiceAgent();
      //                                srAgent.ServiceAgentEmail = attendees[i++].SmtpAddress;
      //                                ServiceAgentWorkingHours srAgntWorkingHours = new ServiceAgentWorkingHours();
      //                                srAgntWorkingHours.StartTime = availability.WorkingHours.StartTime;
      //                                srAgntWorkingHours.EndTime = availability.WorkingHours.EndTime;
      //                                srAgntWorkingHours.ServiceAgentTimeZone = targetTimeZone;
      //                                srAgent.ServiceWorkingHours = srAgntWorkingHours;
      //                                lstTimeSlot = new Collection<TimeSlot>();
      //                                foreach (CalendarEvent calendarItem in availability.CalendarEvents)
      //                                {
      //                                    if (calendarItem.FreeBusyStatus.ToString() != "Free")
      //                                    {
      //                                        TimeSlot tmSlot = new TimeSlot();
      //                                        tmSlot.StartTime = TimeZoneInfo.ConvertTimeFromUtc(calendarItem.StartTime, TimeZoneInfo.FindSystemTimeZoneById(targetTimeZone));
      //                                        tmSlot.EndTime = TimeZoneInfo.ConvertTimeFromUtc(calendarItem.EndTime, TimeZoneInfo.FindSystemTimeZoneById(targetTimeZone));
      //                                        tmSlot.FreeSlot = false;
      //                                        lstTimeSlot.Add(tmSlot);
      //                                    }
      //                                }
      //                                srAgent.BusyTimeSlotsCustomerTimeZone = lstTimeSlot;
      //                                lstServiceAgent.Add(srAgent);
      //                            }
      //                            else
      //                            {
      //                                missedAgentIDs.Add(attendees[i].SmtpAddress);
      //                                i++;
      //                            }
      //                        }

      //                        if (missedAgentIDs.Count > 0 && exchangeRehit == 0)
      //                        {
      //                            attendees.Clear();
      //                            foreach (var agentMissed in missedAgentIDs)
      //                            {
      //                                 attendees.Add(new AttendeeInfo()
      //                                    {
      //                                        SmtpAddress = agentMissed,
      //                                        AttendeeType = MeetingAttendeeType.Required
      //                                    });
      //                            }
      //                            exchangeRehit += 1;
      //                        }
      //                        else if (missedAgentIDs.Count > 0 && exchangeRehit !=0)
      //                        {
      //                            if (exchangeRehit > 2 && missedAgentIDs.Count == 1)
      //                            {
      //                                break;
      //                            }
      //                            else if (exchangeRehit > 2 && missedAgentIDs.Count > 1)
      //                            {
      //                                break;
      //                            }
      //                            attendees.Clear();
      //                            foreach (var agentMissed in missedAgentIDs)
      //                            {
      //                                attendees.Add(new AttendeeInfo()
      //                                {
      //                                    SmtpAddress = agentMissed,
      //                                    AttendeeType = MeetingAttendeeType.Required
      //                                });
      //                            }
      //                            exchangeRehit += 1;
                                  
      //                        }                              
      //                        else if (missedAgentIDs.Count == 0)
      //                        {
      //                            break;
      //                        }
      //                    }
      //                    attendees.Clear();
      //                }

      //            }
      //        }
      //        return lstServiceAgent;
      //    }
      //    catch
      //    {
      //        throw;
      //    }
      //}

      /// <summary>
      /// Gets the user free busy.
      /// </summary>
      /// <param name="agentIDs">The agent I ds.</param>
      /// <param name="targetTimeZone">The target time zone.</param>
      /// <param name="shiftStartTime">The shift start time.</param>
      /// <param name="shiftEndTime">The shift end time.</param>
      /// <param name="startDate">The start date.</param>
      /// <param name="endDate">The end date.</param>
      /// <returns>IReadOnlyList</returns>
      //public IReadOnlyList<ServiceAgent> GetUserFreeBusy(IReadOnlyList<string> agentIDs, string targetTimeZone, TimeSpan shiftStartTime, TimeSpan shiftEndTime,DateTime startDate,DateTime endDate)
      //{
      //    try
      //    {
      //        RequestedTimeWindow = new TimeWindow(startDate, endDate);
      //        List<string> missedAgentIDs = new List<string>();
      //        Collection<TimeSlot> lstTimeSlot;
      //        Collection<ServiceAgent> lstServiceAgent = new Collection<ServiceAgent>();
      //        // Create a list of attendees.
      //        Collection<AttendeeInfo> attendees = new Collection<AttendeeInfo>();
      //        int count = 0;
      //        int totalcount = 0;
      //        if (agentIDs != null || agentIDs.Count > 0)
      //        {
      //            foreach (var agentID in agentIDs)
      //            {
      //                count += 1;
      //                totalcount += 1;
      //                attendees.Add(new AttendeeInfo()
      //                {
      //                    SmtpAddress = agentID,
      //                    AttendeeType = MeetingAttendeeType.Required
      //                });

      //                if ((count == 6) || (totalcount == agentIDs.Count))
      //                {
      //                    count = 0;

      //                    // Specify availability options.
      //                    AvailabilityOptions myOptions = new AvailabilityOptions();
      //                    myOptions.MeetingDuration = 30;
      //                    myOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;



      //                    // Return a set of free/busy times.
      //                    GetUserAvailabilityResults freeBusyResults = this.Exchange.GetUserAvailability(attendees,
      //                                                                                         this.RequestedTimeWindow,
      //                                                                                             AvailabilityData.FreeBusy,
      //                                                                                          myOptions);
      //                    int i = 0;
      //                    foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
      //                    {
      //                        if (availability.WorkingHours != null)
      //                        {
      //                            ServiceAgent srAgent = new ServiceAgent();
      //                            srAgent.ServiceAgentEmail = attendees[i++].SmtpAddress;
      //                            ServiceAgentWorkingHours srAgntWorkingHours = new ServiceAgentWorkingHours();
      //                            srAgntWorkingHours.StartTime = availability.WorkingHours.StartTime;
      //                            srAgntWorkingHours.EndTime = availability.WorkingHours.EndTime;
      //                            srAgntWorkingHours.ServiceAgentTimeZone = targetTimeZone;
      //                            srAgent.ServiceWorkingHours = srAgntWorkingHours;
      //                            lstTimeSlot = new Collection<TimeSlot>();
      //                            foreach (CalendarEvent calendarItem in availability.CalendarEvents)
      //                            {
      //                                if (calendarItem.FreeBusyStatus.ToString() != "Free")
      //                                {
      //                                    TimeSlot tmSlot = new TimeSlot();
      //                                    tmSlot.StartTime = TimeZoneInfo.ConvertTimeFromUtc(calendarItem.StartTime, TimeZoneInfo.FindSystemTimeZoneById(targetTimeZone));
      //                                    tmSlot.EndTime = TimeZoneInfo.ConvertTimeFromUtc(calendarItem.EndTime, TimeZoneInfo.FindSystemTimeZoneById(targetTimeZone));
      //                                    tmSlot.FreeSlot = false;
      //                                    lstTimeSlot.Add(tmSlot);
      //                                }
      //                            }
      //                            srAgent.BusyTimeSlotsCustomerTimeZone = lstTimeSlot;
      //                            lstServiceAgent.Add(srAgent);
      //                        }
      //                        else
      //                        {
      //                            missedAgentIDs.Add(attendees[i++].SmtpAddress);
      //                        }
      //                    }
      //                    attendees.Clear();
      //                }

      //            }
      //        }
      //        return lstServiceAgent;
      //    }
      //    catch
      //    {
      //        throw;
      //    }
      //}

      ///// <summary>
      ///// Saves the appointment.
      ///// </summary>
      ///// <param name="appointmentInfo">The appointment info.</param>
      //public void SaveAppointment(AppointmentInfo appointmentInfo)
      //{
      //    try
      //    {
      //        Appointment appointment = new Appointment(Exchange);
      //        appointment.Subject = appointmentInfo.Subject;
      //        appointment.Body = appointmentInfo.Description;
      //        appointment.Body.BodyType = BodyType.HTML;
      //        appointment.Start = appointmentInfo.Start;
      //        appointment.End = appointmentInfo.End;
      //        appointment.ReminderMinutesBeforeStart = appointmentInfo.Reminder;
      //        appointment.RequiredAttendees.Add(appointmentInfo.AgentMailId, appointmentInfo.AgentMailId);
      //        appointment.RequiredAttendees.Add(appointmentInfo.CustomerMailId, appointmentInfo.CustomerMailId);
      //        appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
      //    }
      //    catch
      //    {
      //        throw;
      //    }
      //}

      /// <summary>
      /// Saves the appointment.
      /// </summary>
      /// <param name="appointmentInfo">The appointment info.</param>
      public void SaveAppointment(AppointmentInfo appointmentInfo)
      {
          try
          {
              // Get the GUID for the property set.
              Guid MyPropertySetId = new Guid("{C11FF724-AA03-4555-9952-8FA248A11C3E}");
              // Create a definition for the extended property.
              ExtendedPropertyDefinition extendedPropertyDefinition =
           new ExtendedPropertyDefinition(MyPropertySetId, "AppointmentID", MapiPropertyType.String);

              Appointment appointment = new Appointment(Exchange);
              appointment.Subject = appointmentInfo.Subject;
              appointment.Body = appointmentInfo.Description;
              appointment.Body.BodyType = BodyType.HTML;
              appointment.Start = appointmentInfo.Start;
              appointment.End = appointmentInfo.End;
              appointment.ReminderMinutesBeforeStart = appointmentInfo.Reminder;
              appointment.RequiredAttendees.Add(appointmentInfo.AgentMailId, appointmentInfo.AgentMailId);
              appointment.RequiredAttendees.Add(appointmentInfo.CustomerMailId, appointmentInfo.CustomerMailId);

              // SetGuidForAppointement(appointment);
              appointment.SetExtendedProperty(extendedPropertyDefinition, appointmentInfo.ID);

              appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
          }
          catch
          {
              throw;
          }
      }

      public void CancelAppointment(AppointmentInfo appointmentInfo)
      {
          try
          {
              // Get the GUID for the property set.
              Guid MyPropertySetId = new Guid("{C11FF724-AA03-4555-9952-8FA248A11C3E}");
              // Create a definition for the extended property.
              ExtendedPropertyDefinition extendedPropertyDefinition = new ExtendedPropertyDefinition(MyPropertySetId, "AppointmentID", MapiPropertyType.String);
              ItemView view = new ItemView(50);
              view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, extendedPropertyDefinition);
              // SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeCreated, DateTime.Today), 
              FindItemsResults<Item> findResults = this.Exchange.FindItems(WellKnownFolderName.Calendar, new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeCreated, appointmentInfo.CancelDate.Date), view);              
              foreach (Appointment item in findResults.Items)
              {                 
                  if (item.ExtendedProperties.Count > 0)
                  {
                      // Display the extended name and value of the extended property.
                      foreach (ExtendedProperty extendedProperty in item.ExtendedProperties)
                      {
                          if (extendedProperty.Value.ToString() == appointmentInfo.ID)
                          {
                              item.CancelMeeting();
                          }
                      }
                  }
              }
          }
          catch
          {
              throw;
          }
      }

      public void UpdateAppointment(AppointmentInfo appointmentInfo)
      {
          try
          {
              // Get the GUID for the property set.
              Guid MyPropertySetId = new Guid("{C11FF724-AA03-4555-9952-8FA248A11C3E}");
              // Create a definition for the extended property.
              ExtendedPropertyDefinition extendedPropertyDefinition = new ExtendedPropertyDefinition(MyPropertySetId, "AppointmentID", MapiPropertyType.String);
              ItemView view = new ItemView(50);
              view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, extendedPropertyDefinition);
              // SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeCreated, DateTime.Today), 
              FindItemsResults<Item> findResults = this.Exchange.FindItems(WellKnownFolderName.Calendar, new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeCreated, appointmentInfo.UpdateDate.Date), view);
              foreach (Appointment item in findResults.Items)
              {
                  if (item.ExtendedProperties.Count > 0)
                  {
                      // Display the extended name and value of the extended property.
                      foreach (ExtendedProperty extendedProperty in item.ExtendedProperties)
                      {
                          if (extendedProperty.Value.ToString() == appointmentInfo.ID)
                          {

                              item.Subject = appointmentInfo.Subject;
                              item.Body = appointmentInfo.Description;
                              item.Body.BodyType = BodyType.HTML;
                              item.Start = appointmentInfo.Start;
                              item.End = appointmentInfo.End;
                              item.ReminderMinutesBeforeStart = appointmentInfo.Reminder;
                              item.RequiredAttendees.Add(appointmentInfo.AgentMailId, appointmentInfo.AgentMailId);
                              item.RequiredAttendees.Add(appointmentInfo.CustomerMailId, appointmentInfo.CustomerMailId);                               
                              item.Update(ConflictResolutionMode.AutoResolve, SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy);
                          }
                      }
                  }
              }
          }
          catch
          {
              throw;
          }
      }

      public void SendAgentUpdateEmail(AppointmentInfo appointmentInfo, string oldlAgentID)
      {
          // Create an email message and identify the Exchange service.
          EmailMessage message = new EmailMessage(Exchange);

          // Add properties to the email message.
          message.Subject = appointmentInfo.Subject; ;
          message.Body = appointmentInfo.Description;
          message.Body.BodyType = BodyType.HTML;          
          message.ToRecipients.Add(oldlAgentID);
          message.CcRecipients.Add(appointmentInfo.AgentMailId);

          // Send the email message and save a copy.
          message.SendAndSaveCopy();


      }
    }
}
