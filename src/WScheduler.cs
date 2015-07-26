//#define debug
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
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using OpenQA.Selenium.PhantomJS;

namespace WScheduler
{
    class Cell { public string Value;}
    class Row { public List<Cell> Cells = new List<Cell>();}

    /// <summary>
    /// Uses Selenium WebDriver and PhantomJS to get the user's work schedule,
    /// parses the html document for the relevant information,
    /// and then upload's the schedule to Google Calendar.
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
        static List<string> m_Schedules = new List<string>();
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "WScheduler";
        static string username = string.Empty, password = string.Empty, success = string.Empty, savedLogin = string.Empty;

        static void Main(string[] args)
        {
            String calendarId = "primary";
            ConsoleColor originalColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Cyan;

            // Setup our Google credentials.
            UserCredential credential = setupGoogleCreds();    
           
            // Get our login info from a file, or the user.
            StreamReader logincreds = null;
            try // Checks the logincreds file to see if the last run was successful,
            {   // and if no input was requested for future runs.
                logincreds = new StreamReader("login.txt");
            }
            catch (FileNotFoundException)
            {
                // don't do anything, if this is the case we'll inform the user in a second.
            }
            finally
            {
                if (logincreds != null)
                {
                    success = logincreds.ReadLine();
                    logincreds.Close();
                }
            }

            try
            {
                logincreds = new StreamReader("login.txt");
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("No Saved Login Credentials Detected.");
            }
            finally
            {
                if (logincreds != null)
                {
                    if (success == "Success")
                    {
                        success = logincreds.ReadLine();
                        username = logincreds.ReadLine();
                        password = logincreds.ReadLine();
                        logincreds.Close();
                    }
                    else
                    {
                        Console.Write("Would you like to use your saved login info? Y/N: ");
                        if (Console.ReadLine().Equals("y", StringComparison.OrdinalIgnoreCase)) // Login creds exist and are used.
                        {
                            username = logincreds.ReadLine();
                            password = logincreds.ReadLine();
                            logincreds.Close();
                        }
                        else // Login creds exist but user doesn't use them.
                        {
                            logincreds.Close();
                            getLoginCreds();
                            saveLoginCreds();
                        }
                    }
                }
                else // Login creds don't exist
                {
                    getLoginCreds();
                    saveLoginCreds();
                }
            }
            

            // GET our schedule
            Console.WriteLine("Getting your schedule... This could take a while...");
            HTTP_GET();

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Parse the DOM, find our tables
#if debug
            Console.WriteLine("Parsing DOM for HTML Tables");
#endif
            List<string> correctTables = new List<string>();
            foreach (string week in m_Schedules)
            {
                correctTables.Add(findTable(week));
            }
#if debug
            Console.WriteLine("Identified which Tables are the Schedule...");
#endif

            // Now we gotta parse our table for the values we want
            List<Row> rows = new List<Row>();
            foreach (string table in correctTables)
            {
                rows.AddRange(parseTable(table));
            }
#if debug
            Console.WriteLine("Parse complete...");
#endif
            Console.WriteLine(string.Empty);
            Console.ForegroundColor = originalColor;
            
           
            // Now that we have all the data in a parsable format
            // we need to parse the "rows" object
            var schedule = parseRows(rows);

            // Display our results to the user.
            displayResults(schedule, originalColor);

            // Now let's upload it to Google Calendar
            Console.WriteLine("Uploading to Google Calendar...");
            int retries = 3;
            while (true)
            {
                try
                {
                    uploadResults(schedule, service, calendarId).Wait();
                    break;
                }
                catch (Exception)
                {
                    if (--retries == 0)
                    {
                        throw;
                    }
                    else
                    Console.WriteLine("Something happened with the upload, let's try again.");
                }
            }
            
            Console.WriteLine("Upload Complete, Press Enter to exit.");

            // After a successful run with input, allow the user to enter "Automate" to remove all needed user input from future runs.
            // Don't prompt the user to use this mode, since it's a hassle to stop automated runs.
            if (success != "Success")
            {
                if ((Console.ReadLine().Equals("Automate") && (savedLogin == "Saved")))
                {
                    successfulRun();
                    Console.WriteLine("Alright, all future runs will require no input.");
                    string loginPath = Path.GetFullPath("login.txt");
                    Console.WriteLine("To reset your password, manually delete the file at: ");
                    Console.WriteLine(string.Empty);
                    Console.WriteLine(loginPath);
                    Console.WriteLine(string.Empty);
                    Console.WriteLine("Press Any key to exit.");
                    Console.ReadKey();
                }
            }
        }

        static public void getLoginCreds()
        {
            Console.Write("Please Enter your username: ");
            username = Console.ReadLine();
            while ((username.Contains("@wegmans.com") || username.Contains("@Wegmans.com")) == false)
            {
                Console.WriteLine("I don't think your username is correct...");
                Console.Write("Please Enter your username: ");
                username = Console.ReadLine();
            }
            Console.Write("Please Enter your password: ");
            password = Console.ReadLine();
        }

        static public void saveLoginCreds()
        {
            Console.Write("Would you like to save your login info? Y/N: ");
            if (Console.ReadLine().Equals("y", StringComparison.OrdinalIgnoreCase))
            {
                string creds = username + "\n" + password;
                File.WriteAllText("login.txt", creds);
                savedLogin = "Saved";
            }
        }

        static public void HTTP_GET()
        {
            var driverService = PhantomJSDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true; // Disables verbose phantomjs output
            IWebDriver driver = new PhantomJSDriver(driverService);
            Console.WriteLine("Logging into Office 365.");
            driver.Navigate().GoToUrl("https://wegmans.sharepoint.com/resources/Pages/LaborPro.aspx");
            if (driver.Title.ToString() == "Sign in to Office 365")
            {
                IWebElement loginentry = driver.FindElement(By.XPath("//*[@id='cred_userid_inputtext']"));
                loginentry.SendKeys(username);
                IWebElement rememberme = driver.FindElement(By.XPath("//*[@id='cred_keep_me_signed_in_checkbox']"));
                rememberme.Click();
            }
            Console.WriteLine("Logging into Sharepoint.");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            try { wait.Until((d) => { return (d.Title.ToString().Contains("Sign In") || d.Title.ToString().Contains("My Schedule")); }); } // Sometimes it skips the second login page.
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Did not recieve an appropriate response from the Sharepoint server. The connection most likely timed out.");
                if (success != "Success")
                {
                    Console.ReadKey();
                }
                driver.Quit();
                Environment.Exit(1);
            }

            if (driver.Title.ToString() == "Sign In")
            {
                IWebElement passwordentry = driver.FindElement(By.XPath("//*[@id='passwordInput']"));
                passwordentry.SendKeys(password);
                passwordentry.Submit();
            }
            try { wait.Until((d) => { return (d.Title.ToString().Contains("Sign In") || d.Title.ToString().Contains("My Schedule")); }); } // Checks to see if the password was incorrect.
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Did not recieve an appropriate response from the Sharepoint server. The connection most likely timed out.");
                if (success != "Success")
                {
                    Console.ReadKey();
                }
                driver.Quit();
                Environment.Exit(1);
            }
            if (driver.Title.ToString() == "Sign In")
            {

                IWebElement error = driver.FindElement(By.XPath("//*[@id='error']"));
                var errorString = error.Text.ToString();
                if (errorString.Contains("Incorrect user ID or password"))
                {
                    while (driver.Title.ToString() == "Sign In")
                    {
                        IWebElement usernameentry = driver.FindElement(By.XPath("//*[@id='userNameInput']"));
                        IWebElement passwordentry = driver.FindElement(By.XPath("//*[@id='passwordInput']"));
                        usernameentry.Clear();
                        passwordentry.Clear();
                        Console.WriteLine("You seem to have entered the wrong username or password.");
                        getLoginCreds();
                        usernameentry.SendKeys(username);
                        passwordentry.SendKeys(password);
                        passwordentry.Submit();
                    }
                    saveLoginCreds();
                }
                else
                {
                    Console.WriteLine("An unexpected error has occured with the webpage.");
                    Console.WriteLine(errorString);
                    driver.Quit();
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }

            Console.WriteLine("Waiting for LaborPro...");
            try { wait.Until((d) => { return (d.SwitchTo().Frame(0)); }); } // Waits for the inline frame to load.
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("LaborPro link's inline frame was not generated properly.");
                driver.Quit();
                if (success != "Success")
                {
                    Console.ReadKey();
                }
                Environment.Exit(1);
            }
            string BaseWindow = driver.CurrentWindowHandle;
            try { wait.Until((d) => { return (d.FindElement(By.XPath("/html/body/a"))); }); } // Waits until javascript generates the SSO link.
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("LaborPro SSO Link was not generated properly.");
                if (success != "Success")
                {
                    Console.ReadKey();
                }
                driver.Quit();
                Environment.Exit(1);
            }
            IWebElement accessschedule = driver.FindElement(By.XPath("/html/body/a"));
            accessschedule.Click();
            string popupHandle = string.Empty;
            ReadOnlyCollection<string> windowHandles = driver.WindowHandles;

            foreach (string handle in windowHandles)
            {
                if (handle != driver.CurrentWindowHandle)
                {
                    popupHandle = handle;
                    break;
                }
            }
            driver.SwitchTo().Window(popupHandle);

            Console.WriteLine("Accessing LaborPro.");
            try { wait.Until((d) => { return (d.Title.ToString().Contains("Welcome")); }); }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Did not properly switch to LabroPro Window.");
                if (success != "Success")
                {
                    Console.ReadKey();
                }
                driver.Quit();
                Environment.Exit(1);
            }
            m_Schedules.Add(driver.PageSource.ToString());
            for (int i = 0; i < 2; i++) // Clicks "Next" and gets the schedules for the next two weeks.
            {
                driver.FindElement(By.XPath("//*[@id='pageBody']/form/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/table[1]/tbody/tr/td[1]/a[3]")).Click();
                m_Schedules.Add(driver.PageSource.ToString());
            }

            driver.Quit();
            Console.WriteLine("Got your Schedule.");
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
                Console.WriteLine("Google Credentials file saved to: " + credPath);
                return credential;
            }
        }

        static public string findTable(string file)
        {
            var endTable = "</table>";
            var startTable = "<table";
            int index = 0;
            m_Tables.Clear();

            // Parse the DOM, find our tables
            while ((index = file.IndexOf(startTable, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                var tableContentEndIndex = file.IndexOf(endTable, index, StringComparison.OrdinalIgnoreCase);
                var tableContent = file.Substring(index, tableContentEndIndex - index + endTable.Length);

#if debug
                Console.WriteLine("Found a table DOM element!");
#endif

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

#if debug
                Console.WriteLine("Found a row within the table...");
#endif

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

#if debug
                    Console.WriteLine("Found a cell within the row with value '{0}'", actualValue);
#endif
                    row.Cells.Add(new Cell { Value = actualValue });
                }

                rows.Add(row);
            }

            return rows;
        }

        static public List<WorkDay> parseRows(List<Row> rows)
        {
            var schedule = new List<WorkDay>();
            for (int i = 0; i < 3; i++)
            {
                var dateRow = rows[1 + (i * 8)]; // First row is empty, then adds a week each iteration.
                var scheduleHoursRow = rows[2 + (i * 8)];
                var activityRow = rows[3 + (i * 8)];
                var locationRow = rows[4 + (i * 8)];
                var scheduleTimesRow = rows[5 + (i * 8)];
                var commentsRow = rows[6 + (i * 8)];

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
            }
            return schedule;
        }

        static public void displayResults(List<WorkDay> schedule, ConsoleColor originalColor)
        {
            foreach (var day in schedule)
            {
                if (day.Hours == 0)
                {
                    //Console.WriteLine("You get " + day.Day.ToString() + " off!");
                    //Console.WriteLine(string.Empty);
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

        static public async Task uploadResults(List<WorkDay> schedule, CalendarService service, string calendarId)
        {
            foreach (var day in schedule)
            {
                var request = new BatchRequest(service);
                // Setup request for current events.
                EventsResource.ListRequest listrequest = service.Events.List(calendarId);
                listrequest.TimeMin = DateTime.Today.AddDays(-6);
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
        
        static public void successfulRun()
        {
            string creds = "Success\n" + username + "\n" + password;
            File.WriteAllText("login.txt", creds);
        }

    }
}