namespace Today
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Outlook;

    internal static class Program
    {
        internal static void Main(string[] args)
        {
            try
            {
                DateTime start = ParseCommandLine(args);
                DateTime end = start.AddDays(1);
                PrintAppointments(start, end);
            }
            catch (System.Exception x)
            {
                Console.WriteLine("Error: {0}", x.Message);
            }
        }

        private static DateTime ParseCommandLine(string[] args)
        {
            if (args.Length == 0)
            { // Just today's appointments.
                return DateTime.Today;
            }
            else if (args.Length == 1)
            { // Appointments for the given day, as indicated by the offset.
                int offset;
                if (!int.TryParse(args[0], out offset))
                {
                    PrintUsage();
                    throw new ArgumentException(string.Format("{0} is not a valid integer", args[0]));
                }

                return DateTime.Today.AddDays(offset);
            }
            else
            {
                PrintUsage();
                throw new ArgumentException("Wrong number of arguments");
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("Today\t<-- Print today's appointments.");
            Console.WriteLine("Today n\t<-- Print appointments for today+n days (n may be negative).");
        }

        private static void PrintAppointments(DateTime start, DateTime end)
        {
            var outlook = new Application();
            var calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            Items appointments = calendar.Items;
            appointments.IncludeRecurrences = true;
            appointments.Sort("[Start]");
            const string DateFormat = "M/dd/yy h:mm tt";
            string filter = string.Format("[Start]>=\"{0}\" AND [Start]<\"{1}\"", start.ToString(DateFormat), end.ToString(DateFormat));
            var todaysAppointments = appointments.Restrict(filter);
            var appointment = todaysAppointments.GetFirst();
            while (appointment != null)
            {
                Console.Write("{0}-{1} {2}", appointment.Start.ToShortTimeString(), appointment.End.ToShortTimeString(), appointment.Subject);
                if (!string.IsNullOrEmpty(appointment.Location))
                {
                    Console.Write(" ({0})", appointment.Location);
                }

                Console.WriteLine();

                appointment = todaysAppointments.GetNext();
            }
        }
    }
}
