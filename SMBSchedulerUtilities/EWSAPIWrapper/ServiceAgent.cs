// --------------------------------------------------------------------------
// <copyright file="ServiceAgent.cs" company="Microsoft Corporation">
//   Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------
namespace EWSAPIWrapper
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Exchange.WebServices.Data;
    using System.Collections.ObjectModel;    

    /// <summary>
    /// This class represents the free/busy time slots & other information related to a Service Agent.
    /// </summary>
  public class ServiceAgent
    {
        /// <summary>
        /// Gets or sets the service agent email.
        /// </summary>
        /// <value>
        /// The service agent email.
        /// </value>
        public string ServiceAgentEmail { get; set; }

        /// <summary>
        /// Gets or sets the free time slots.
        /// </summary>
        /// <value>
        /// The free time slots.
        /// </value>
        public Collection<TimeSlot> FreeBusyTimeSlots { get; set; }

        /// <summary>
        /// Gets or sets the free time slots customer time zone.
        /// </summary>
        /// <value>
        /// The free time slots customer time zone.
        /// </value>
        public Collection<TimeSlot> FreeTimeSlotsCustomerTimeZone { get; set; }

        /// <summary>
        /// Gets or sets the free time slots.
        /// </summary>
        /// <value>
        /// The free time slots.
        /// </value>
        public Collection<TimeSlot> FreeBusyTimeSlotsCustomerTimeZone { get; set; }

        /// <summary>
        /// Gets or sets the busy time slots.
        /// </summary>
        /// <value>
        /// The busy time slots.
        /// </value>
        public Collection<TimeSlot> BusyTimeSlots { get; set; }

        /// <summary>
        /// Gets or sets the busy time slots customer time zone.
        /// </summary>
        /// <value>
        /// The busy time slots customer time zone.
        /// </value>
        public Collection<TimeSlot> BusyTimeSlotsCustomerTimeZone { get; set; }

        /// <summary>
        /// Gets or sets the service working hours.
        /// </summary>
        /// <value>
        /// The service working hours.
        /// </value>
        public ServiceAgentWorkingHours ServiceWorkingHours { get; set; }

        /// <summary>
        /// Gets or sets the meeting time slot.
        /// </summary>
        /// <value>
        /// The meeting time slot.
        /// </value>
        public int MeetingTimeSlot { get; set; }

        /// <summary>
        /// Converts the Service Agent time to the customer time based on the customer time zones.
        /// </summary>
        /// <param name="timeZone">The time zone.</param>
        public void ConvertTimeBetweenTheTimeZones(TimeZoneInfo timeZone)
        {
            try
            {
                List<TimeSlot> freeBusyTimeSlotsCustomerTimeZone = new List<TimeSlot>();
                foreach (var tmSlot in this.FreeBusyTimeSlots)
                {
                    TimeSlot timeSlot = new TimeSlot();
                    timeSlot.StartTime = TimeZoneInfo.ConvertTime(tmSlot.StartTime, TimeZoneInfo.FindSystemTimeZoneById(this.ServiceWorkingHours.ServiceAgentTimeZone), timeZone);
                    timeSlot.EndTime = TimeZoneInfo.ConvertTime(tmSlot.EndTime, TimeZoneInfo.FindSystemTimeZoneById(this.ServiceWorkingHours.ServiceAgentTimeZone), timeZone);
                    timeSlot.FreeSlot = tmSlot.FreeSlot;
                    freeBusyTimeSlotsCustomerTimeZone.Add(timeSlot);
                }

                this.FreeTimeSlotsCustomerTimeZone = new Collection<TimeSlot>((from x in freeBusyTimeSlotsCustomerTimeZone
                                                                               where x.FreeSlot.Equals(true)
                                                                               select x).ToList());
                this.BusyTimeSlots = new Collection<TimeSlot>((from x in freeBusyTimeSlotsCustomerTimeZone
                                                               where x.FreeSlot.Equals(false)
                                                               select x).ToList());
                IList<TimeSlot> iListfreeBusyTimeSlotsCustomerTimeZone = freeBusyTimeSlotsCustomerTimeZone;
                this.FreeBusyTimeSlotsCustomerTimeZone = new Collection<TimeSlot>(iListfreeBusyTimeSlotsCustomerTimeZone);
            }
            catch
            {
                throw;
            }
        }
      
        /// <summary>
        /// Creates the service agent free busy time slots.
        /// </summary>
        /// <param name="timeWindow">The time window.</param>
        /// <param name="meetingTimeSlot">The meeting time slot.</param>
        public void CreateServiceAgentFreeBusyTimeSlots(DateTime startDate, DateTime endDate, int meetingTimeSlot, TimeSpan shiftStartTime, TimeSpan shiftEndTime)
        {
            try
            {
                DateTime startDateTime = new DateTime();
                startDateTime = startDate;
                this.FreeBusyTimeSlots = new Collection<TimeSlot>();
                while (DateTime.Compare(startDateTime.Date, endDate.Date) != 0)
                {
                    ////Create free time slots based on minimum meeting time slots per team
                    ////specified by input parameter meetingTimeSlot
                    TimeSpan startTimeSpan = shiftStartTime;
                    TimeSpan endTimeSpan = shiftEndTime;
                    TimeSpan agentStartTimeSpan = this.ServiceWorkingHours.StartTime;
                    TimeSpan agentEndTimeSpan = this.ServiceWorkingHours.EndTime;
                    Collection<TimeSlot> freeTimeSlots = new Collection<TimeSlot>();
                    while (startTimeSpan.CompareTo(endTimeSpan) != 0)
                    {
                        TimeSlot tmSlot = new TimeSlot();
                        tmSlot.StartTime = new DateTime(startDateTime.Year, startDateTime.Month, startDateTime.Day, startTimeSpan.Hours, startTimeSpan.Minutes, startTimeSpan.Seconds);
                        TimeSpan tmpTimespan = startTimeSpan.Add(TimeSpan.FromMinutes(meetingTimeSlot));
                        startTimeSpan = tmpTimespan;
                        tmSlot.EndTime = new DateTime(startDateTime.Year, startDateTime.Month, startDateTime.Day, startTimeSpan.Hours, startTimeSpan.Minutes, startTimeSpan.Seconds);
                        tmSlot.FreeSlot = true;
                        TimeSpan freeSlotStartTimeSpan = new TimeSpan(tmSlot.StartTime.Hour, tmSlot.StartTime.Minute, tmSlot.StartTime.Second);
                        TimeSpan freeSlotEndTimeSpan = new TimeSpan(tmSlot.EndTime.Hour, tmSlot.EndTime.Minute, tmSlot.EndTime.Second);
                        //Mark the slot busy(tmSlot.FreeSlot = false) if it is out of range of agent's working hours. 
                        if ((freeSlotStartTimeSpan < agentStartTimeSpan) || (freeSlotEndTimeSpan > agentEndTimeSpan))
                        {
                            tmSlot.FreeSlot = false;
                        }
                        else
                            tmSlot.FreeSlot = true;


                        freeTimeSlots.Add(tmSlot);
                    }

                    // Marking the busy time slots based on the busy meeting hours returned by EWS
                    List<TimeSlot> lstBusyTimeSlot = new List<TimeSlot>();

                    //lstBusyTimeSlot = (from x in this.BusyTimeSlotsCustomerTimeZone
                    //                   where (DateTime.Compare(x.StartTime.Date, startDateTime.Date) == 0) &&
                    //                   (x.StartTime.Hour >= this.ServiceWorkingHours.StartTime.Hours && x.StartTime.Hour <= this.ServiceWorkingHours.EndTime.Hours && (x.StartTime != x.EndTime))
                    //                   select x).ToList<TimeSlot>();
                    lstBusyTimeSlot = (from x in this.BusyTimeSlotsCustomerTimeZone
                                       where (DateTime.Compare(x.StartTime.Date, startDateTime.Date) == 0) ||( (x.StartTime.Date < startDateTime.Date) && (x.EndTime.Date > startDateTime.Date))                            
                                       select x).ToList<TimeSlot>();
                    //stBusyTimeSlot = this.BusyTimeSlotsCustomerTimeZone.ToList<TimeSlot>(); ;
                    foreach (TimeSlot busySlot in lstBusyTimeSlot)
                    {
                        bool isBusyInterval = false;
                        foreach (TimeSlot freeSlot in freeTimeSlots)
                        {
                            TimeSpan freeSlotStartTimeSpan = new TimeSpan(freeSlot.StartTime.Hour, freeSlot.StartTime.Minute, freeSlot.StartTime.Second);
                            TimeSpan freeSlotEndTimeSpan = new TimeSpan(freeSlot.EndTime.Hour, freeSlot.EndTime.Minute, freeSlot.EndTime.Second);
                            if ((freeSlotStartTimeSpan < agentStartTimeSpan) || (freeSlotEndTimeSpan > agentEndTimeSpan))
                            {
                                freeSlot.FreeSlot = false;
                                
                            }
                            else if ((busySlot.StartTime <= freeSlot.StartTime && busySlot.EndTime >= freeSlot.EndTime))
                            {
                                freeSlot.FreeSlot = false;
                                if (busySlot.EndTime == freeSlot.EndTime)
                                {
                                break;
                                }
                                
                            }
                            else if ((busySlot.StartTime == freeSlot.StartTime && busySlot.EndTime < freeSlot.EndTime))
                            {
                                freeSlot.FreeSlot = false;
                                break;

                            }
                            else if ((busySlot.StartTime < freeSlot.StartTime) && (busySlot.EndTime > freeSlot.StartTime && busySlot.EndTime < freeSlot.EndTime))
                            {
                                freeSlot.FreeSlot = false;                                
                                break;                               
                                
                            }
                            else if ((busySlot.StartTime > freeSlot.StartTime && busySlot.StartTime < freeSlot.EndTime))
                            {
                                freeSlot.FreeSlot = false;
                                if (busySlot.EndTime == freeSlot.EndTime)
                                {
                                    break;
                                }

                            }
                            
                            
                            //else if ((busySlot.StartTime >= freeSlot.StartTime && busySlot.EndTime <= freeSlot.EndTime) && isBusyInterval.Equals(false))
                            //{
                            //    freeSlot.FreeSlot = false;
                            //    break;
                            //}
                            //else if ((busySlot.StartTime >= freeSlot.StartTime && busySlot.StartTime < freeSlot.EndTime) && isBusyInterval.Equals(false))
                            //{
                            //    freeSlot.FreeSlot = false;
                            //    isBusyInterval = true;
                            //    if (busySlot.EndTime <= freeSlot.EndTime)
                            //    {
                            //        break;
                            //    }
                            //}
                            //else if ((busySlot.StartTime >= freeSlot.StartTime && busySlot.EndTime <= freeSlot.EndTime) && isBusyInterval.Equals(false))
                            //{
                            //    freeSlot.FreeSlot = false;
                            //    break;
                            //}
                            //else if ((busySlot.StartTime >= freeSlot.StartTime && busySlot.StartTime < freeSlot.EndTime) && isBusyInterval.Equals(false))
                            //{
                            //    freeSlot.FreeSlot = false;
                            //    isBusyInterval = true;
                            //    if (busySlot.EndTime <= freeSlot.EndTime)
                            //    {
                            //        break;
                            //    }
                            //}
                            //else if ((busySlot.EndTime > freeSlot.StartTime && busySlot.EndTime <= freeSlot.EndTime) && isBusyInterval.Equals(true))
                            //{
                            //    freeSlot.FreeSlot = false;
                            //    break;
                            //}
                            //else if ((busySlot.EndTime > freeSlot.StartTime && busySlot.EndTime > freeSlot.EndTime) && isBusyInterval.Equals(true))
                            //{
                            //    freeSlot.FreeSlot = false;
                            //}
                        }
                    }

                    if (this.FreeBusyTimeSlots != null)
                    {
                        IEnumerable<TimeSlot> combinedCollection = this.FreeBusyTimeSlots.Concat(freeTimeSlots);
                        this.FreeBusyTimeSlots = new Collection<TimeSlot>(combinedCollection.ToList());
                    }
                    else
                    {
                        this.FreeBusyTimeSlots = freeTimeSlots;
                    }

                    startDateTime = startDateTime.AddDays(1);
                }
            }
            catch
            {
                throw;
            }
        }


        public void CreateServiceAgentMeetingSlots(int meetingTimeSlot)
        {
            List<TimeSlot> freeBusyTimeSlotsCustomerTimeZone = new List<TimeSlot>();
            bool freeslot = false;
            int count = 0;
            TimeSlot prevTimeSlot = new TimeSlot();
            int slotCount = 0;
            foreach (var tmSlot in this.FreeBusyTimeSlots)
            {
                slotCount++;
                if (slotCount == 9 && meetingTimeSlot == 60)
                {
                    tmSlot.FreeSlot = false;
                    slotCount = 0;
                    continue;
                }

                if (tmSlot.FreeSlot == true && count == 0)
                {
                    count = 1;
                    freeslot = true;
                    prevTimeSlot = tmSlot;
                }
                else if (tmSlot.FreeSlot == false && count == 0)
                {
                    count = 1;
                    freeslot = false;
                }
                else if ((tmSlot.FreeSlot == true || tmSlot.FreeSlot == false) && count == 1)
                {
                    if (freeslot == true && tmSlot.FreeSlot == true)
                    {
                        tmSlot.FreeSlot = false;
                        count = 0;
                    }
                    else if (freeslot == true && tmSlot.FreeSlot == false)
                    {
                        prevTimeSlot.FreeSlot = false;
                        count = 0;
                    }
                    else if (freeslot == false)
                    {
                        tmSlot.FreeSlot = false;
                        count = 0;
                    }

                }
            }
        }


   
    }
}
