//#define debug
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Net.Http;
using System.Security.Cryptography;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Requests;

using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Firefox;


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
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "WScheduler";

        static void Main(string[] args)
        {
            String calendarId = "primary";
            ConsoleColor originalColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Cyan;

            // Setup our Google credentials.
            UserCredential credential = SetupGoogleCreds();    
           
            // Get our login info from a file, or the user.
            string username = string.Empty, password = string.Empty;
            string automate = string.Empty, savedlogin = string.Empty;
            savedlogin = LoadLoginCreds(ref username, ref password); 

            // GET our schedule
            Console.WriteLine("Getting your schedule... This could take a while...");
            List<string> m_Schedules = new List<string>();
            int retries = 3;
            int errcode = 0;
            while (retries > 0)
            {
                errcode = HTTP_GET(username, password, automate, savedlogin, m_Schedules);
                if (errcode == 0)
                {
                    break;
                }
                else
                {
                    Console.WriteLine("There was a problem getting your schedule, Error Code:" + errcode);
                    Console.WriteLine("Let's try loggin in again.");
                }
                retries--;
                if (retries == 0)
                {
                    Console.WriteLine("We tried getting your schedule three times and it didn't work.");
                    Console.WriteLine("Something must be up with your connection to the server.");
                    Console.WriteLine("Please check your connection and try again.");
                    if (automate != "Automate")
                    {
                        Console.ReadKey();
                    }
                    Environment.Exit(0);
                }
            }

            // Create Google Calendar API service.
            CalendarService service = new CalendarService(new BaseClientService.Initializer()
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
                correctTables.Add(FindTable(week));
            }
#if debug
            Console.WriteLine("Identified which Tables are the Schedule...");
#endif

            // Now we gotta parse our table for the values we want
            List<Row> rows = new List<Row>();
            foreach (string table in correctTables)
            {
                rows.AddRange(ParseTable(table));
            }
#if debug
            Console.WriteLine("Parse complete...");
#endif
            Console.WriteLine(string.Empty);
            Console.ForegroundColor = originalColor;
            
           
            // Now that we have all the data in a parsable format
            // we need to parse the "rows" object
            List<WorkDay> schedule = ParseRows(rows);

            // Display our results to the user.
            DisplayResults(schedule, originalColor);

            // Now let's upload it to Google Calendar
            Console.WriteLine("Uploading to Google Calendar...");
            retries = 3;
            while (true)
            {
                try
                {
                    UploadResults(schedule, service, calendarId).Wait();
                    break;
                }
                catch (Exception ex)
                {
                    if (--retries == 0)
                    {
                        Console.WriteLine("We tried uploading your schedule three times and it didn't work.");
                        Console.WriteLine("Here's a nasty error message to explain why:");
                        Console.Write(ex.ToString());
                    }
                    else
                    Console.WriteLine("Something happened with the upload, let's try again.");
                }
            }
            
            Console.WriteLine("Upload Complete, Press Enter to exit.");

            //After a successful run with input, allow the user to enter "Automate" to remove all needed user input from future runs.
            //Don't prompt the user to use this mode, since it's a hassle to stop automated runs.
            if (automate != "Automate")
            {
                if ((Console.ReadLine().Equals("Automate") && (savedlogin == "Saved")))
                {
                    //AutomateRun(username, password);
                    //Console.WriteLine("Alright, all future runs will require no input.");
                    //string loginPath = Path.GetFullPath("login.txt");
                    //Console.WriteLine("To reset your password, manually delete the file at: ");
                    //Console.WriteLine(string.Empty);
                    //Console.WriteLine(loginPath);
                    //Console.WriteLine(string.Empty);
                    //Console.WriteLine("Press Any key to exit.");
                    //Console.ReadKey();
                }
            }
        }

        static public string LoadLoginCreds(ref string username, ref string password)
        {
            string savedlogin = string.Empty;
            string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            string logincredPath = Path.Combine(credPath, ".credentials/login.txt");
            string entropyPath = Path.Combine(credPath, ".credentials/entropy.txt");
            byte[] logincreds = null, entropy = null;
            string ptcreds = string.Empty;
            string[] creds;
            try
            {
                logincreds = File.ReadAllBytes(logincredPath);
                entropy = File.ReadAllBytes(entropyPath);
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("No Saved Login Credentials Detected.");
            }
            finally
            {
                if (logincreds != null)
                {
                    Console.Write("Would you like to use your saved login info? Y/N: ");
                    if (Console.ReadLine().Equals("y", StringComparison.OrdinalIgnoreCase)) // Login creds exist and are used.
                    {
                        byte[] credbytes = ProtectedData.Unprotect(logincreds, entropy, DataProtectionScope.CurrentUser);
                        ptcreds = Encoding.UTF8.GetString(credbytes);
                        creds = ptcreds.Split('\n');
                        username = creds[0];
                        password = creds[1];

                    }
                    else // Login creds exist but user doesn't use them.
                    {
                        GetLoginCreds(ref username, ref password);
                        SaveLoginCreds(username, password, ref savedlogin);
                    }

                }
                else // Login creds don't exist
                {
                    GetLoginCreds(ref username, ref password);
                    SaveLoginCreds(username, password, ref savedlogin);
                }
            }
            return savedlogin;
        }

        static public void GetLoginCreds(ref string username, ref string password)
        {
            Console.Write("Please Enter your username: ");
            username = Console.ReadLine();
            while ((username.Contains("@wegmans.com") || username.Contains("@Wegmans.com")) == false)
            {
                Console.WriteLine("I don't think your username is correct...");
                Console.Write("Please Enter your username@wegmans.com: ");
                username = Console.ReadLine();
            }
            Console.Write("Please Enter your password: ");
            const int ENTER = 13, BACKSP = 8, CTRLBACKSP = 127;
            int[] FILTERED = { 0, 27, 9, 10 /*, 32 space, if you care */ };
            Stack<char> pass = new Stack<char>();
            char chr = (char)0;

            while ((chr = System.Console.ReadKey(true).KeyChar) != ENTER)
            {
                if (chr == BACKSP)
                {
                    if (pass.Count > 0)
                    {
                        System.Console.Write("\b \b");
                        pass.Pop();
                    }
                }
                else if (chr == CTRLBACKSP)
                {
                    while (pass.Count > 0)
                    {
                        System.Console.Write("\b \b");
                        pass.Pop();
                    }
                }
                else if (FILTERED.Count(x => chr == x) > 0) { }
                else
                {
                    pass.Push((char)chr);
                    System.Console.Write("*");
                }
            }
            Console.WriteLine();

            // Popping the password off the stack will result in a reverse password,
            // So let's flip it.
            string revpassword = string.Join("", pass.ToArray());
            char[] revArray = revpassword.ToCharArray();
            Array.Reverse(revArray);
            password = new string(revArray);
        }

        static public void SaveLoginCreds(string username, string password, ref string savedlogin)
        {
            string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
            string logincredPath = Path.Combine(credPath, ".credentials/login.txt");
            string entropyPath = Path.Combine(credPath, ".credentials/entropy.txt");
            Console.Write("Would you like to save your login info? Y/N: ");

            if (Console.ReadLine().Equals("y", StringComparison.OrdinalIgnoreCase))
            {
                string logincreds = username + "\n" + password;
                byte[] plaintextcreds = Encoding.UTF8.GetBytes(logincreds);
                byte[] entropy = new byte[20];
                using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider())
                {
                    rng.GetBytes(entropy);
                }

                byte[] encryptedcreds = ProtectedData.Protect(plaintextcreds, entropy, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(logincredPath, encryptedcreds);
                File.WriteAllBytes(entropyPath, entropy);
                savedlogin = "Saved";
            }
        }

        static public int HTTP_GET(string username, string password, string savedlogin, string automate, List<string> m_Schedules)
        {
            var driverService = PhantomJSDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true; // Disables verbose phantomjs output
            IWebDriver driver = new PhantomJSDriver(driverService);
            //IWebDriver driver = new FirefoxDriver(); // Debug with firefox.

            Console.WriteLine("Logging into Office 365.");
            driver.Navigate().GoToUrl("https://wegmans.sharepoint.com/resources/Pages/LaborPro.aspx");
            if (driver.Title.ToString() == "Sign in to Office 365")
            {
                IWebElement loginentry = driver.FindElement(By.XPath("//*[@id='cred_userid_inputtext']"));
                loginentry.SendKeys(username);
                IWebElement rememberme = driver.FindElement(By.XPath("//*[@id='cred_keep_me_signed_in_checkbox']"));
                rememberme.Click();
            }

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            try { wait.Until((d) => { return (d.Title.ToString().Contains("Sign In") || d.Title.ToString().Contains("My Schedule")); }); } // Sometimes it skips the second login page.
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Did not recieve an appropriate response from the Sharepoint server. The connection most likely timed out.");
                driver.Quit();
                return 1;
            }
            Console.WriteLine("Logging into Sharepoint.");
            
            if (driver.Title.ToString() == "Sign In")
            {
                try { wait.Until((d) => { return (d.FindElement(By.XPath("//*[@id='passwordInput']"))); }); }
                catch (Exception)
                {
                    Console.WriteLine("Password input box did not load correctly.");
                    driver.Quit();
                    return 2;
                }
                IWebElement passwordentry = driver.FindElement(By.XPath("//*[@id='passwordInput']"));
                passwordentry.SendKeys(password);
                passwordentry.Submit();
            }

            try { wait.Until((d) => { return (d.Title.ToString().Contains("Sign In") || d.Title.ToString().Contains("My Schedule")); }); } // Checks to see if the password was incorrect.
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Did not recieve an appropriate response from the Sharepoint server. The connection most likely timed out.");
                driver.Quit();
                return 3;
            }
            if (driver.Title.ToString() == "Sign In")
            {

                IWebElement error = driver.FindElement(By.XPath("//*[@id='error']"));
                string errorString = error.Text.ToString();
                if (errorString.Contains("Incorrect user ID or password"))
                {
                    while (driver.Title.ToString() == "Sign In")
                    {
                        IWebElement usernameentry = driver.FindElement(By.XPath("//*[@id='userNameInput']"));
                        IWebElement passwordentry = driver.FindElement(By.XPath("//*[@id='passwordInput']"));
                        usernameentry.Clear();
                        passwordentry.Clear();
                        Console.WriteLine("You seem to have entered the wrong username or password.");
                        GetLoginCreds(ref username, ref password);
                        Console.WriteLine("Trying again...");
                        usernameentry.SendKeys(username);
                        passwordentry.SendKeys(password);
                        passwordentry.Submit();
                    }
                    SaveLoginCreds(username, password, ref savedlogin);
                }
                else
                {
                    Console.WriteLine("An unexpected error has occured with the webpage.");
                    Console.WriteLine(errorString);
                    driver.Quit();
                    return 4;
                }
            }

            Console.WriteLine("Waiting for LaborPro...");
            int retries = 2;
            while (true) // Retry because this error can be solved by a simple page reload.
            {
                try { wait.Until((d) => { return (d.SwitchTo().Frame(0)); }); break; } // Waits for the inline frame to load.
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("LaborPro link's inline frame was not generated properly.");
                    Console.WriteLine("Reloading the page...");
                    driver.Navigate().Refresh();
                    retries--;
                    if (retries <= 0)
                    {
                        driver.Quit();
                        return 5;
                    }

                }
            }
            
            string BaseWindow = driver.CurrentWindowHandle;
            try { wait.Until((d) => { return (d.FindElement(By.XPath("/html/body/a"))); }); } // Waits until javascript generates the SSO link.
            catch (Exception)
            {
                Console.WriteLine("LaborPro SSO Link was not generated properly.");
                Console.WriteLine("You encountered the classic \"SadPage\" error. We need to try logging in again...");
                driver.Quit();
                return 6;
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
                return 7;
            }
            m_Schedules.Add(driver.PageSource.ToString());
            for (int i = 0; i < 2; i++) // Clicks "Next" and gets the schedules for the next two weeks.
            {
                driver.FindElement(By.XPath("//*[@id='pageBody']/form/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/table[1]/tbody/tr/td[1]/a[3]")).Click();
                m_Schedules.Add(driver.PageSource.ToString());
            }

            driver.Quit();
            Console.WriteLine("Got your Schedule.");
            return 0;
        }

        static public UserCredential SetupGoogleCreds()
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
                Console.WriteLine("Credentials files saved to: " + credPath);
                return credential;
            }
        }

        static public string FindTable(string file)
        {
            string endTable = "</table>";
            string startTable = "<table";
            int index = 0;
            List<string> m_Tables = new List<string>();

            // Parse the DOM, find our tables
            while ((index = file.IndexOf(startTable, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                int tableContentEndIndex = file.IndexOf(endTable, index, StringComparison.OrdinalIgnoreCase);
                string tableContent = file.Substring(index, tableContentEndIndex - index + endTable.Length);

#if debug
                Console.WriteLine("Found a table DOM element!");
#endif

                m_Tables.Add(tableContent);
            }

            // Identify the table that's actually relevant to us
            string keyword = "Schedule Times:";
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

        static public List<Row> ParseTable(string correctTable)
        {
            const string rowStart = "<tr";
            const string cellStart = "<td";
            const string rowEnd = "</tr>";
            const string cellEnd = "</td>";

            var rows = new List<Row>();

            int index = 0;
            while ((index = correctTable.IndexOf(rowStart, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                int rowContentEndIndex = correctTable.IndexOf(rowEnd, index);
                String rowContent = correctTable.Substring(index, rowContentEndIndex - index + rowEnd.Length);
                var row = new Row();

#if debug
                Console.WriteLine("Found a row within the table...");
#endif

                int index2 = 0;
                while ((index2 = rowContent.IndexOf(cellStart, index2 + 1, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    int cellContentEndIndex = rowContent.IndexOf(cellEnd, index2);
                    String cellContent = rowContent.Substring(index2, cellContentEndIndex - index2 + cellEnd.Length);

                    // We have to parse the cell html for the value
                    // we just want the value between ...>HERE<...
                    int startOpen = cellContent.IndexOf(">");
                    int endOpen = cellContent.IndexOf("<", cellStart.Length); // start past the <td> part
                    String actualValue = cellContent
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

        static public List<WorkDay> ParseRows(List<Row> rows)
        {
            var schedule = new List<WorkDay>();
            for (int i = 0; i < 3; i++)
            {
                Row dateRow = rows[1 + (i * 8)]; // First row is empty, then adds a week each iteration.
                Row scheduleHoursRow = rows[2 + (i * 8)];
                Row activityRow = rows[3 + (i * 8)];
                Row locationRow = rows[4 + (i * 8)];
                Row scheduleTimesRow = rows[5 + (i * 8)];
                Row commentsRow = rows[6 + (i * 8)];

                int dayIndex = 1; // 0 is column names
                while (true)
                {
                    var workDay = new WorkDay();
                    String dateString = dateRow.Cells[dayIndex].Value;

                    // Get the day of the week
                    workDay.Day = WorkDay.Get(dateString);
                    if (workDay.Day == WorkDay.DayEnum.INVALID)
                    {
                        break;
                    }

                    // split the string "Tue 6/23"
                    // and use the second part ['Tue', '6/23']
                    String datePart = dateString.Split(' ')[1];
                    workDay.Date = DateTime.Parse(datePart);

                    workDay.Hours = float.Parse(scheduleHoursRow.Cells[dayIndex].Value.Replace(":", "."));
                    workDay.Activity = activityRow.Cells[dayIndex].Value;
                    workDay.Comments = commentsRow.Cells[dayIndex].Value;
                    workDay.Location = locationRow.Cells[dayIndex].Value;

                    String timeSpanPart = scheduleTimesRow.Cells[dayIndex].Value;
                    String[] times = timeSpanPart.Split('-'); // split '2:00 AM-8:00 PM' on the '-'

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

        static public void DisplayResults(List<WorkDay> schedule, ConsoleColor originalColor)
        {
            foreach (WorkDay day in schedule)
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

        static public async Task UploadResults(List<WorkDay> schedule, CalendarService service, string calendarId)
        {
            foreach (WorkDay day in schedule)
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
                String workevent = "Wegmans: ";
                Events eventslist = listrequest.Execute();
                if (eventslist.Items != null && eventslist.Items.Count > 0)
                {
                    foreach (Event eventItem in eventslist.Items)
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

        static public void AutomateRun(string username, string password) // todo: update to proper credential storage.
        {
            string creds = "Automate\n" + username + "\n" + password;
            File.WriteAllText("login.txt", creds);
        }

    }
}