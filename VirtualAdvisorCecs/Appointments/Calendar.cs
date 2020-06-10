using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System.Configuration;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2.Flows;

using AdaptiveCards;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace VirtualAdvisorCecs.Appointments
{
    public class Calendar
    {
        private static string ApplicationName = ConfigurationManager.AppSettings["ApplicationName"].ToString();
        private static string ClientId = ConfigurationManager.AppSettings["ClientId"].ToString();
        private static string ClientSecret = ConfigurationManager.AppSettings["ClientSecret"].ToString();
        private static string RedirectURL = ConfigurationManager.AppSettings["RedirectURL"].ToString();

        private static ClientSecrets GoogleClientSecrets = new ClientSecrets()
        {
            ClientId = ClientId,
            ClientSecret = ClientSecret
        };

        public static string[] Scopes =
        {
                                CalendarService.Scope.Calendar,
                                CalendarService.Scope.CalendarReadonly
                            };

        public static UserCredential GetUserCredential(out string error)
        {
            UserCredential credential = null;
            error = string.Empty;

            try
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = ClientId,
                    ClientSecret = ClientSecret
                },
                Scopes,
                "cecscaptain.umich@gmail.com",//"svaddire@umich.edu",/*"hemanth.vaddireddy8@gmail.com",*///Environment.UserName,
                CancellationToken.None,
                null).Result;
            }
            catch (Exception ex)
            {
                credential = null;
                error = "Failed to UserCredential Initialization: " + ex.ToString();
            }

            return credential;
        }
        public static IAuthorizationCodeFlow GoogleAuthorizationCodeFlow(out string error)
        {
            IAuthorizationCodeFlow flow = null;
            error = string.Empty;

            try
            {
                flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = GoogleClientSecrets,
                    Scopes = Scopes
                });
            }
            catch (Exception ex)
            {
                flow = null;
                error = "Failed to AuthorizationCodeFlow Initialization: " + ex.ToString();
            }

            return flow;
        }

        public static UserCredential GetGoogleUserCredentialByRefreshToken(string refreshToken, out string error)
        {
            Google.Apis.Auth.OAuth2.Responses.TokenResponse respnseToken = null;
            UserCredential credential = null;
            string flowError;
            error = string.Empty;
            try
            {
                // Get a new IAuthorizationCodeFlow instance
                IAuthorizationCodeFlow flow = GoogleAuthorizationCodeFlow(out flowError);

                respnseToken = new Google.Apis.Auth.OAuth2.Responses.TokenResponse() { RefreshToken = refreshToken };

                // Get a new Credential instance                
                if ((flow != null && string.IsNullOrWhiteSpace(flowError)) && respnseToken != null)
                {
                    credential = new UserCredential(flow, "user", respnseToken);
                }

                // Get a new Token instance
                if (credential != null)
                {
                    bool success = credential.RefreshTokenAsync(CancellationToken.None).Result;
                }

                // Set the new Token instance
                if (credential.Token != null)
                {
                    string newRefreshToken = credential.Token.RefreshToken;
                }
            }
            catch (Exception ex)
            {
                credential = null;
                error = "UserCredential failed: " + ex.ToString();
            }
            return credential;
        }

        public static CalendarService GetCalendarService(string refreshToken, out string error)
        {
            CalendarService calendarService = null;
            string credentialError;
            error = string.Empty;
            try
            {
                var credential = GetGoogleUserCredentialByRefreshToken(refreshToken, out credentialError);
                if (credential != null && string.IsNullOrWhiteSpace(credentialError))
                {
                    calendarService = new CalendarService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName
                    });
                }
            }
            catch (Exception ex)
            {
                calendarService = null;
                error = "Calendar service failed: " + ex.ToString();
            }
            return calendarService;
        }

        public static string AddCalenderEvents(string refreshToken, string emailAddress, string summary, DateTime start, string Student, string Advisor, out string error)
        {
            string eventId = string.Empty;
            error = string.Empty;
            string serviceError;

            try
            {
                var calendarService = GetCalendarService(refreshToken, out serviceError);

                if (calendarService != null && string.IsNullOrWhiteSpace(serviceError))
                {
                    var list = calendarService.CalendarList.List().Execute();
                    var calendar = list.Items.SingleOrDefault(c => c.Summary == emailAddress);
                    if (calendar != null)
                    {
                        Google.Apis.Calendar.v3.Data.Event calenderEvent = new Google.Apis.Calendar.v3.Data.Event();

                        calenderEvent.Summary = "Appointment";
                        calenderEvent.Description = "Appointment with Advisor";
                        calenderEvent.Location = "CIS 210";
                        calenderEvent.Start = new Google.Apis.Calendar.v3.Data.EventDateTime
                        {
                            //DateTime = new DateTime(2018, 1, 20, 19, 00, 0)
                            DateTime = start//,
                                            //TimeZone = "Europe/Istanbul"
                        };
                        calenderEvent.End = new Google.Apis.Calendar.v3.Data.EventDateTime
                        {
                            //DateTime = new DateTime(2018, 4, 30, 23, 59, 0)
                            DateTime = start.AddHours(0.5)//,
                                                          //TimeZone = "Europe/Istanbul"
                        };
                        calenderEvent.Recurrence = new List<string>();

                        //Set Remainder
                        calenderEvent.Reminders = new Google.Apis.Calendar.v3.Data.Event.RemindersData()
                        {
                            UseDefault = false,
                            Overrides = new Google.Apis.Calendar.v3.Data.EventReminder[]
                            {
                                                    new Google.Apis.Calendar.v3.Data.EventReminder() { Method = "email", Minutes = 24 * 60 },
                                                    new Google.Apis.Calendar.v3.Data.EventReminder() { Method = "popup", Minutes = 24 * 60 }
                            }
                        };

                        #region Attendees
                        //Set Attendees
                        calenderEvent.Attendees = new Google.Apis.Calendar.v3.Data.EventAttendee[] {
                                                new Google.Apis.Calendar.v3.Data.EventAttendee() { Email = Student /*"hemanth260292@gmail.com"*/ },
                                                new Google.Apis.Calendar.v3.Data.EventAttendee() { Email = Advisor /*"hemanth.vaddireddy8@gmail.com"*/ }
                                            };
                        #endregion

                        var newEventRequest = calendarService.Events.Insert(calenderEvent, calendar.Id);
                        newEventRequest.SendNotifications = true;
                        var eventResult = newEventRequest.Execute();
                        eventId = eventResult.Id;
                    }
                }
            }
            catch (Exception ex)
            {
                eventId = string.Empty;
                error = ex.Message;
            }
            return eventId;
        }
        public static Google.Apis.Calendar.v3.Data.Event UpdateCalenderEvents(string refreshToken, string emailAddress, string summary, DateTime? start, DateTime? end, string eventId, out string error)
        {
            Google.Apis.Calendar.v3.Data.Event eventResult = null;
            error = string.Empty;
            string serviceError;
            try
            {
                var calendarService = GetCalendarService(refreshToken, out serviceError);
                if (calendarService != null)
                {
                    var list = calendarService.CalendarList.List().Execute();
                    var calendar = list.Items.SingleOrDefault(c => c.Summary == emailAddress);
                    if (calendar != null)
                    {
                        // Define parameters of request
                        EventsResource.ListRequest request = calendarService.Events.List("primary");
                        request.TimeMin = DateTime.Now;
                        request.ShowDeleted = false;
                        request.SingleEvents = true;
                        request.MaxResults = 10;
                        request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                        // Get selected event
                        Google.Apis.Calendar.v3.Data.Events events = request.Execute();
                        var selectedEvent = events.Items.FirstOrDefault(c => c.Id == eventId);
                        if (selectedEvent != null)
                        {
                            selectedEvent.Summary = summary;
                            selectedEvent.Start = new Google.Apis.Calendar.v3.Data.EventDateTime
                            {
                                DateTime = start
                            };
                            selectedEvent.End = new Google.Apis.Calendar.v3.Data.EventDateTime
                            {
                                DateTime = start.Value.AddHours(12)
                            };
                            selectedEvent.Recurrence = new List<string>();

                            // Set Remainder
                            selectedEvent.Reminders = new Google.Apis.Calendar.v3.Data.Event.RemindersData()
                            {
                                UseDefault = false,
                                Overrides = new Google.Apis.Calendar.v3.Data.EventReminder[]
                                {
                                                        new Google.Apis.Calendar.v3.Data.EventReminder() { Method = "email", Minutes = 24 * 60 },
                                                        new Google.Apis.Calendar.v3.Data.EventReminder() { Method = "popup", Minutes = 24 * 60 }
                                }
                            };

                            // Set Attendees
                            selectedEvent.Attendees = new Google.Apis.Calendar.v3.Data.EventAttendee[]
                            {
                                                    new Google.Apis.Calendar.v3.Data.EventAttendee() { Email = "svaddire@umich.edu" },
                                                    new Google.Apis.Calendar.v3.Data.EventAttendee() { Email = emailAddress }
                            };
                        }

                        var updateEventRequest = calendarService.Events.Update(selectedEvent, calendar.Id, eventId);
                        updateEventRequest.SendNotifications = true;
                        eventResult = updateEventRequest.Execute();
                    }
                }
            }
            catch (Exception ex)
            {
                eventResult = null;
                error = ex.ToString();
            }
            return eventResult;
        }
        public static void DeleteCalendarEvents(string refreshToken, string emailAddress, string eventId, out string error)
        {
            string result = string.Empty;
            error = string.Empty;
            string serviceError;
            try
            {
                var calendarService = GetCalendarService(refreshToken, out serviceError);
                if (calendarService != null)
                {
                    var list = calendarService.CalendarList.List().Execute();
                    var calendar = list.Items.FirstOrDefault(c => c.Summary == emailAddress);
                    if (calendar != null)
                    {
                        // Define parameters of request
                        EventsResource.ListRequest request = calendarService.Events.List("primary");
                        request.TimeMin = DateTime.Now;
                        request.ShowDeleted = false;
                        request.SingleEvents = true;
                        request.MaxResults = 10;
                        request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                        // Get selected event
                        Google.Apis.Calendar.v3.Data.Events events = request.Execute();
                        var selectedEvent = events.Items.FirstOrDefault(c => c.Id == eventId);
                        if (selectedEvent != null)
                        {
                            var deleteEventRequest = calendarService.Events.Delete(calendar.Id, eventId);
                            deleteEventRequest.SendNotifications = true;
                            deleteEventRequest.Execute();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result = string.Empty;
                error = ex.ToString();
            }
        }

        public string Run(JObject timeSlot, string studentInfo, string advisorInfo)
        {
            string refreshToken = string.Empty;
            string credentialError;
            var credential = GetUserCredential(out credentialError);
            if (credential != null && string.IsNullOrWhiteSpace(credentialError))
            {
                //Save RefreshToken into Database 
                refreshToken = credential.Token.RefreshToken;
            }

            string addEventError;
            string calendarEventId = string.Empty;

            var startTime = timeSlot["TimeSlot"].ToString();
            var jsonObjectStudent = (JObject)JsonConvert.DeserializeObject(studentInfo);
            var Date = Convert.ToDateTime(jsonObjectStudent["Date"].ToString());
            var appointmentStartDateTime = Date.Add(TimeSpan.Parse(startTime));

            var Student = jsonObjectStudent["Email"].ToString();

            calendarEventId = AddCalenderEvents(refreshToken, "hemanthv@umich.edu", "My Calendar Event", appointmentStartDateTime, Student, advisorInfo, out addEventError);
            return startTime;

            //string updateEventError;
            //if (!string.IsNullOrEmpty(calendarEventId))
            //{
            //    UpdateCalenderEvents(refreshToken, "svaddire@umich.edu", "Modified Calendar Event ", DateTime.Now, DateTime.Now.AddDays(1), calendarEventId, out updateEventError);
            //}

            //string deleteEventError;
            //if (!string.IsNullOrEmpty(calendarEventId))
            //{
            //    DeletCalendarEvents(refreshToken, "svaddire@umich.edu", calendarEventId, out deleteEventError);
            //}
        }
        public string FixAppointment(string startTimeSlot, string appointmentDate, string studentEmailID, string advisorEmailID)
        {
            string refreshToken = string.Empty;
            string credentialError;
            var credential = GetUserCredential(out credentialError);
            if (credential != null && string.IsNullOrWhiteSpace(credentialError))
            {
                //Save RefreshToken into Database 
                refreshToken = credential.Token.RefreshToken;
            }

            string addEventError;
            string calendarEventId = string.Empty;

            var Date = Convert.ToDateTime(appointmentDate);
            var appointmentStartDateTime = Convert.ToDateTime(startTimeSlot);

            calendarEventId = AddCalenderEvents(refreshToken, "cecscaptain.umich@gmail.com", "My Calendar Event", appointmentStartDateTime, studentEmailID, advisorEmailID, out addEventError);
            return startTimeSlot;
        }
    }
}