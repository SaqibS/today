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
                var outlook = new Application();
                var calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                Items appointments = calendar.Items;
                appointments.IncludeRecurrences = true;
                appointments.Sort("[Start]");
                DateTime today = DateTime.Today, tomorrow = today.AddDays(1);
                const string DateFormat = "M/dd/yy h:mm tt";
                string filter = string.Format("[Start]>=\"{0}\" AND [Start]<\"{1}\"", today.ToString(DateFormat), tomorrow.ToString(DateFormat));
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
            catch (System.Exception x)
            {
                Console.WriteLine("Error: {0}", x.Message);
            }
        }
    }
}
