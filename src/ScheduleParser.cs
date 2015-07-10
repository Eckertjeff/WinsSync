using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Requests;
using System.Threading;
using System.Net.Http;

namespace ScheduleParser
{
    class Cell { public string Value;}
    class Row { public List<Cell> Cells = new List<Cell>(); }

    /// <summary>
    /// Represents a single record based on an HTML table in the DOM.
    /// </summary>
    class WorkDay
    {
        public DayEnum Day;
        public DateTime Date;
        public float Hours;
        public string Activity;
        public string Location;
        public string StartTime;
        public string EndTime;
        public DateTime StartDateTime;
        public DateTime EndDateTime;
        public string Comments;

        public enum DayEnum { Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, INVALID }
        public static DayEnum Get(string value)
        {
            if (value.Contains("Sun"))
            {
                return DayEnum.Sunday;
            }
            else if (value.Contains("Mon"))
            {
                return DayEnum.Monday;
            }
            else if (value.Contains("Tue"))
            {
                return DayEnum.Tuesday;
            }
            else if (value.Contains("Wed"))
            {
                return DayEnum.Wednesday;
            }
            else if (value.Contains("Thu"))
            {
                return DayEnum.Thursday;
            }
            else if (value.Contains("Fri"))
            {
                return DayEnum.Friday;
            }
            else if (value.Contains("Sat"))
            {
                return DayEnum.Saturday;
            }

            return DayEnum.INVALID;
        }

    }

    class Program
    {
        static List<string> m_Tables = new List<string>();
        static string file2 = string.Empty;
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "WScheduler";

        static void Main(string[] args)
        {
            String calendarId = "primary";
            ConsoleColor originalColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Cyan;
            var file = File.ReadAllText("Welcome.html");

            // GET our schedule
            Task t = new Task(HTTP_GET);
            t.Start();
            
            // Setup our Google credentials.
            UserCredential credential = setupGoogleCreds();

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Parse the DOM, find our tables
            Console.WriteLine("Parsing DOM for HTML Tables");
            var correctTable = findTable(file);
            Console.WriteLine("Identified which Table is the Schedule...");
            

            // Now we gotta parse our table for the values we want
            var rows = parseTable(correctTable);
            Console.WriteLine("Parse complete...");
            Console.WriteLine(string.Empty);
            Console.ForegroundColor = originalColor;

            // Now that we have all the data in a parsable format
            // we need to parse the "rows" object
            var schedule = parseRows(rows);

            // Display our results to the user.
            displayResults(schedule, originalColor);

            // Now let's upload it to Google Calendar
            Console.WriteLine("Uploading to Google Calendar...");
            uploadResults(schedule, service, calendarId);

            Console.WriteLine("Upload Complete, Press any key to exit.");
            Console.ReadKey();
        }

        static public async void HTTP_GET()
        {
            var TARGETURL = "https://schedule.mywegmansconnect.com/wfm/action/silentSignon?loginInfo=IjkmIFOw9odOxz9SU5PcX0eOZrqY%2bR9zqj6Xmfamhlc%3d";

            HttpClientHandler handler = new HttpClientHandler()
                {
                    //put cookiesncredzzzz here
                };

            Console.WriteLine("GET: + " + TARGETURL);

            // Use HttpClient.            
            HttpClient client = new HttpClient(handler);

            HttpResponseMessage response = await client.GetAsync(TARGETURL);
            HttpContent content = response.Content;

            // Check Status Code                                
            Console.WriteLine("Response StatusCode: " + (int)response.StatusCode);
            // Read the string.
            file2 = await content.ReadAsStringAsync();

            // Display the result.
            Console.WriteLine(file2);
        }
       
        static public UserCredential setupGoogleCreds()
        {
            UserCredential credential;
            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
                return credential;
            }
        }

        static public string findTable(string file)
        {
            var endTable = "</table>";
            var startTable = "<table";
            int index = 0;

            // Parse the DOM, find our tables
            while ((index = file.IndexOf(startTable, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                var tableContentEndIndex = file.IndexOf(endTable, index, StringComparison.OrdinalIgnoreCase);
                var tableContent = file.Substring(index, tableContentEndIndex - index + endTable.Length);

                Console.WriteLine("Found a table DOM element!");

                m_Tables.Add(tableContent);
            }

            // Identify the table that's actually relevant to us
            var keyword = "Schedule Times:";
            var oneWeWantList = m_Tables.Where(s => s.Contains(keyword));

            if (oneWeWantList.Count() > 1)
            {
                throw new Exception("Found more than one 'right' table.");
            }

            var correctTable = oneWeWantList.FirstOrDefault();

            if (correctTable == null)
            {
                throw new Exception("The 'right' table does not exist.");
            }

            return correctTable;
        }

        static public List<Row> parseTable(string correctTable)
        {
            const string rowStart = "<tr";
            const string cellStart = "<td";
            const string rowEnd = "</tr>";
            const string cellEnd = "</td>";

            var rows = new List<Row>();

            var index = 0;
            while ((index = correctTable.IndexOf(rowStart, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                var rowContentEndIndex = correctTable.IndexOf(rowEnd, index);
                var rowContent = correctTable.Substring(index, rowContentEndIndex - index + rowEnd.Length);
                var row = new Row();

                Console.WriteLine("Found a row within the table...");

                var index2 = 0;
                while ((index2 = rowContent.IndexOf(cellStart, index2 + 1, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    var cellContentEndIndex = rowContent.IndexOf(cellEnd, index2);
                    var cellContent = rowContent.Substring(index2, cellContentEndIndex - index2 + cellEnd.Length);

                    // We have to parse the cell html for the value
                    // we just want the value between ...>HERE<...
                    var startOpen = cellContent.IndexOf(">");
                    var endOpen = cellContent.IndexOf("<", cellStart.Length); // start past the <td> part
                    var actualValue = cellContent
                        .Substring(startOpen + 1, endOpen - startOpen - 1)
                        .Replace("\n", string.Empty)
                        .Replace("\r", string.Empty)
                        .Replace("&nbsp;", string.Empty)
                        .Trim();

                    if (actualValue.Equals("&nbsp;"))
                    {
                        actualValue = string.Empty; // Non breaking space is space...
                    }

                    Console.WriteLine("Found a cell within the row with value '{0}'", actualValue);
                    row.Cells.Add(new Cell { Value = actualValue });
                }

                rows.Add(row);
            }

            return rows;
        }

        static public List<WorkDay> parseRows(List<Row> rows)
        {
            var schedule = new List<WorkDay>();

            var dateRow = rows[1]; // First row is empty... classic
            var scheduleHoursRow = rows[2];
            var activityRow = rows[3];
            var locationRow = rows[4];
            var scheduleTimesRow = rows[5];
            var commentsRow = rows[6];

            int dayIndex = 1; // 0 is column names
            while (true)
            {
                var workDay = new WorkDay();
                var dateString = dateRow.Cells[dayIndex].Value;

                // Get the day of the week
                workDay.Day = WorkDay.Get(dateString);
                if (workDay.Day == WorkDay.DayEnum.INVALID)
                {
                    break;
                }

                // split the string "Tue 6/23"
                // and use the second part ['Tue', '6/23']
                var datePart = dateString.Split(' ')[1];
                workDay.Date = DateTime.Parse(datePart);

                workDay.Hours = float.Parse(scheduleHoursRow.Cells[dayIndex].Value.Replace(":", "."));
                workDay.Activity = activityRow.Cells[dayIndex].Value;
                workDay.Comments = commentsRow.Cells[dayIndex].Value;
                workDay.Location = locationRow.Cells[dayIndex].Value;

                var timeSpanPart = scheduleTimesRow.Cells[dayIndex].Value;
                var times = timeSpanPart.Split('-'); // split '2:00 AM-8:00 PM' on the '-'

                workDay.StartTime = times[0];
                workDay.EndTime = times.Length == 1 ? times[0] : times[1];

                if (workDay.Hours > 0)
                {
                    workDay.StartDateTime = DateTime.Parse(workDay.Date.ToShortDateString() + " " + workDay.StartTime);
                    workDay.EndDateTime = DateTime.Parse(workDay.Date.ToShortDateString() + " " + workDay.EndTime);
                }

                schedule.Add(workDay);
                dayIndex++;
            }
            return schedule;
        }

        static public void displayResults(List<WorkDay> schedule, ConsoleColor originalColor)
        {
            foreach (var day in schedule)
            {
                if (day.Hours == 0)
                {
                    Console.WriteLine("You get " + day.Day.ToString() + " off!");
                    Console.WriteLine(string.Empty);
                    continue;
                }

                if (day.Hours > 8)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("******** WARNING - ANOMALY DETECTED: ********");
                }

                Console.WriteLine(day.Date.ToShortDateString() + " at " + day.Location + Environment.NewLine + " from " +
                    day.StartTime + " to " + day.EndTime + Environment.NewLine + " doing " + day.Activity
                    + " (" + day.Comments + ") total of " + day.Hours + " hours");
                Console.WriteLine(string.Empty);
                Console.ForegroundColor = originalColor;
            }
        }

        static public async void uploadResults(List<WorkDay> schedule, CalendarService service, string calendarId)
        {
            foreach (var day in schedule)
            {
                var request = new BatchRequest(service);
                // Setup request for current events.
                EventsResource.ListRequest listrequest = service.Events.List(calendarId);
                listrequest.TimeMin = DateTime.Now;
                listrequest.TimeMax = DateTime.Today.AddDays(15);
                listrequest.ShowDeleted = false;
                listrequest.SingleEvents = true;
                listrequest.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                // Check to see if work events are already in place on the schedule, if they are,
                // setup a batch request to delete them.
                // Not the best implementation, but takes care of duplicates without tracking
                // eventIds to update existing events.
                var workevent = "Wegmans: ";
                Events eventslist = listrequest.Execute();
                if (eventslist.Items != null && eventslist.Items.Count > 0)
                {
                    foreach (var eventItem in eventslist.Items)
                    {
                        DateTime eventcontainer = (DateTime)eventItem.Start.DateTime; // Typecast to use ToShortDateString() method for comparison.
                        if (((eventcontainer.ToShortDateString()) == (day.Date.ToShortDateString())) && (eventItem.Summary.Contains(workevent)))
                        {
                            request.Queue<Event>(service.Events.Delete(calendarId, eventItem.Id),
                                (content, error, i, message) =>
                                {
                                    if (error != null)
                                    {
                                        throw new Exception(error.ToString());
                                    }
                                });
                        }
                    }
                }
                // Setup a batch request to upload the work events.
                request.Queue<Event>(service.Events.Insert(
                new Event
                {
                    Summary = workevent + day.Activity,
                    Description = day.Comments,
                    Location = day.Location,
                    Start = new EventDateTime()
                    {
                        DateTime = day.StartDateTime,
                        TimeZone = "America/New_York",
                    },
                    End = new EventDateTime()
                    {
                        DateTime = day.EndDateTime,
                        TimeZone = "America/New_York",
                    },
                    Reminders = new Event.RemindersData()
                    {
                        UseDefault = true,
                    },
                }, calendarId),
                (content, error, i, message) =>
                {
                    if (error != null)
                    {
                        throw new Exception(error.ToString());
                    }
                });
                // Execute batch request.
                await request.ExecuteAsync();
            }
        }

    }
}