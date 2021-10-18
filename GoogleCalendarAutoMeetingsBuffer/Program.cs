using System;

namespace GoogleCalendarAutoMeetingsBuffer
{
    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Calendar.v3;
    using Google.Apis.Calendar.v3.Data;
    using Google.Apis.Services;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    namespace CalendarQuickstart
    {
        class Program
        {
            static void Main(string[] args)
            {
                string jsonFile = "responsive-hall-328817-9ab1e4575be3.json";
                string calendarToAddEventsId = @"wmatecki97@gmail.com";
                string calendarToReadEventsId = @"wiktor.matecki@mtab.com";
                string eventsTimeZone = "Europe/Tirane";
                int bufferDuration = 5;

                string[] Scopes = { CalendarService.Scope.Calendar };
                CalendarService service = GetCalendarService(jsonFile, Scopes);

                // Define parameters of request.
                EventsResource.ListRequest listRequest = service.Events.List(calendarToReadEventsId);
                listRequest.TimeMin = DateTime.Now;
                listRequest.ShowDeleted = false;
                listRequest.SingleEvents = true;
                listRequest.MaxResults = 10;
                listRequest.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                // List events.
                Events events = listRequest.Execute();

                var calendarEvents = events.Items.OrderBy(i => i.Start.DateTime).ToList();
                for (int i = 1; i < calendarEvents.Count - 1; i++)
                {
                    var previousevent = calendarEvents[i - 1];
                    var currentEvent = calendarEvents[i];
                    var nextEvent = calendarEvents[i + 1];

                    if (currentEvent.Start.TimeZone == eventsTimeZone && (currentEvent.End.DateTime.Value.Subtract(currentEvent.Start.DateTime.Value)).TotalMinutes == bufferDuration)
                    {
                        if (previousevent.End.DateTime < currentEvent.Start.DateTime && nextEvent.Start.DateTime > currentEvent.End.DateTime)
                        {
                            service.Events.Delete(calendarToAddEventsId, currentEvent.Id).Execute();
                        }
                        continue;
                    }

                    if (previousevent.End.DateTime < currentEvent.Start.DateTime)
                    {
                        var startDate = currentEvent.Start.DateTime.Value;
                        startDate = startDate.AddMinutes(-bufferDuration);
                        var endDate = currentEvent.Start.DateTime.Value;
                        AddEvent(calendarToAddEventsId, service, startDate, endDate, calendarToReadEventsId);
                    }

                    if (nextEvent.Start.DateTime > currentEvent.End.DateTime)
                    {
                        var startDate = currentEvent.End.DateTime.Value;
                        var endDate = currentEvent.End.DateTime.Value;
                        endDate = endDate.AddMinutes(bufferDuration);
                        AddEvent(calendarToAddEventsId, service, startDate, endDate, calendarToReadEventsId);
                    }
                }
            }

            private static CalendarService GetCalendarService(string jsonFile, string[] Scopes)
            {
                ServiceAccountCredential credential;

                using (var stream =
                    new FileStream(jsonFile, FileMode.Open, FileAccess.Read))
                {
                    var confg = Google.Apis.Json.NewtonsoftJsonSerializer.Instance.Deserialize<JsonCredentialParameters>(stream);
                    credential = new ServiceAccountCredential(
                       new ServiceAccountCredential.Initializer(confg.ClientEmail)
                       {
                           Scopes = Scopes
                       }.FromPrivateKey(confg.PrivateKey));
                }

                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Calendar API Sample",
                });
                return service;
            }

            private static void AddEvent(string calendarId, CalendarService service, DateTime startDate, DateTime endDate, string atendeeEmail)
            {
                Event bufferEvent = new Event()
                {
                    Summary = "Buffer time",
                    Description = "Buffer time between the meetings",
                    Start = new EventDateTime()
                    {
                        DateTime = startDate
                    },
                    End = new EventDateTime()
                    {
                        DateTime = endDate,
                    },
                    Attendees = new List<EventAttendee>
                    {
                        new EventAttendee(){Email = atendeeEmail}
                    }
                };
                var insertRequest = service.Events.Insert(bufferEvent, calendarId);
                insertRequest.SendNotifications = false;
                insertRequest.Execute();
            }
        }
    }
}
