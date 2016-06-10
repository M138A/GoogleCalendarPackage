using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CalendarModule
{
    public class Calendar
    {
        private Events events = null;
        private string[] Scopes = new string[] {
            CalendarService.Scope.Calendar, // Manage your calendars
 	        CalendarService.Scope.CalendarReadonly // View your Calendars
        };
        private string ApplicationName = "TamTam reservation service";
        CalendarService service = null;
        public Calendar()
        {
            setup();
        }
        /*
        This method creates the Google Calendar Service
        */
        private void setup()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.ReadWrite))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar-auth.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }            
            service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        /*
        Search for an event by name. Case-sensitive!
        This method will return all the events, when you only need one event use searchEventByName
        */

        public List<Event> searchEventsByName(string name)
        {
            List<Event> l = new List<Event>();
            foreach (var eventItem in events.Items)
            {
                if (eventItem.Summary.Contains(name))
                {

                    l.Add(eventItem);
                }
            }
            return l;
        }
        /*
        Search for an event by name. Case-sensitive!
        This method will only return 1 event, for more events use searchEventsByName
        */

        public Event searchEventByName(string name)
        {
            foreach (var eventItem in events.Items)
            {

                if (eventItem.Summary.Contains(name))
                {

                    return eventItem;
                }

            }
            return null;
        }
        /*
        Gets the newest version of the calendar
        */
        public void updateEvents()
        {
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            events = request.Execute();

        }
        /*
       Returns all the events in the calendar.
       */
        public Events getEvents(){ return events; }


        /*
        Prints a list to the console with a list of all the events in the Calendar
        */
        public void printEventsList()
        {
            Console.WriteLine("Upcoming events:");
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    Console.WriteLine("{0} ({1})", eventItem.Summary, when);
                }
            }
            else
            {
                Console.WriteLine("No upcoming events found.");
            }
            Console.Read();
        }
        /*
        Method that adds an event to the Google Calendar
        Summary - string : Enter a summary of the event (Title)
        Location - string : Enter the location of the event, if there's no location provide an empty string
        Description - string : Enter a description of the event, if there's no description provide an empty string
        startDateTime - string : Enter the datetime when the event starts in a format parsable by DateTime
        endDateTime - string : Enter the datetime when the event ends in a format parsable by DateTime
        */


        public void createEvent(string summary, string location, string description, string startDateTime, string endDateTime)
        {

            Event newEvent = new Event()
            {
                Summary = summary,
                Location = location,
                Description = description,
                Start = new EventDateTime()
                {
                    DateTime = DateTime.Parse(startDateTime),
                    TimeZone = "Europe/Amsterdam",
                },
                End = new EventDateTime()
                {
                    DateTime = DateTime.Parse(endDateTime),
                    TimeZone = "Europe/Amsterdam",
                },
            };
            String calendarId = "primary";
            EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
            Event createdEvent = request.Execute();
        }
    }

}
