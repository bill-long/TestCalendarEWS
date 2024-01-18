using Microsoft.Exchange.WebServices.Data;
using System;

namespace bilong.TestCalendarEWS
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: {0} <server> <mailbox>", Environment.GetCommandLineArgs()[0]);
                return;
            }

            var server = args[0];
            var mailbox = args[1];
            var verbose = args.Length > 2;

            var exchService = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Credentials = System.Net.CredentialCache.DefaultNetworkCredentials,
                Url = new Uri($"https://{server}/EWS/Exchange.asmx")
            };

            var mbx = new Mailbox(mailbox);
            var calendarId = new FolderId(WellKnownFolderName.Calendar, mbx);
            var arrayOfPropertiesToLoad = new[]{
                AppointmentSchema.End,
                AppointmentSchema.EndTimeZone,
                AppointmentSchema.IsAllDayEvent,
                AppointmentSchema.IsOnlineMeeting,
                AppointmentSchema.Location,
                AppointmentSchema.MyResponseType,
                AppointmentSchema.Organizer,
                AppointmentSchema.Start,
                AppointmentSchema.StartTimeZone,
                AppointmentSchema.AppointmentType,
                AppointmentSchema.IsResponseRequested,
                ItemSchema.IsReminderSet,
                ItemSchema.Subject,
                ItemSchema.Sensitivity,
                ItemSchema.LastModifiedTime
            };
            var propertySet = new PropertySet(arrayOfPropertiesToLoad);
            var calendarView = new CalendarView(DateTime.Now.AddMonths(-2), DateTime.Now.AddMonths(2), 500);
            calendarView.PropertySet = propertySet;

            var appointments = exchService.FindAppointments(calendarId, calendarView);
            Console.WriteLine($"Found {appointments.Items.Count} appointments");
            Console.WriteLine();

            foreach (var appointment in appointments)
            {
                var isAllDayEvent = false;
                try
                {
                    isAllDayEvent = appointment.IsAllDayEvent;
                    if (verbose)
                    {
                        Console.WriteLine($"UniqueId: {appointment.Id.UniqueId}");
                        Console.WriteLine($"Subject: {appointment.Subject}");
                        Console.WriteLine($"IsAllDayEvent: {isAllDayEvent}");
                        Console.WriteLine();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"UniqueId: {appointment.Id.UniqueId}");
                    Console.WriteLine($"Subject: {appointment.Subject}");
                    Console.WriteLine($"IsAllDayEvent: {ex.Message}");
                    Console.WriteLine();
                }
            }
        }
    }
}
