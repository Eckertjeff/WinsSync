using System;
using System.Security.Cryptography;

namespace WinsSyncConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            WinsSync winssync = new WinsSync();

            //Setup our credentials.
            winssync.SetupGoogleCreds();
            try
            {
                winssync.LoadLoginCreds();
            }
            catch (CryptographicException)
            {
                // Saved credentials are somehow corrupt, let's just delete them and get them from the user.
                winssync.DeleteLoginCreds();
                winssync.LoadLoginCreds();
            }

            // Time to log in and get our schedule.
            // Retry loop in case javascript execution breaks. It happens often enough, and we need to log out, and back in to recover.
            int retries = 5;
            while (retries > 0)
            {
                var errorcode = winssync.HTTP_GET();
                if(errorcode == 0)
                {
                    break;
                }
                if(retries == 0)
                {
                    Console.WriteLine("We tried getting your schedule five times and it didn't work.");
                    Console.WriteLine("Something must be up with your connection to the server.");
                    Console.WriteLine("Please check your connection and try again. Press any key to quit.");
                    if (winssync.Automate != "Automate")
                    {
                        Console.ReadKey();
                    }
                    Environment.Exit(0);
                }
                retries--;
            }

            // Parse the HTML for our schedule info.
            winssync.FindTable();
            winssync.ParseTable();
            winssync.ParseRows();

            // Let's let the user look at their schedule while we're busy uploading.
            winssync.DisplayResults();

            // Now we need to upload our schedule to Google.
            // Sometimes the upload request isn't processed correctly, just retrying fixes it almost every time.
            Console.WriteLine("Uploading your schedule now...");
            retries = 5;
            while (true)
            {
                try
                {
                    winssync.UploadResults().Wait();
                    break;
                }
                catch (Exception ex)
                {
                    if (--retries == 0)
                    {
                        Console.WriteLine("We tried uploading your schedule multiple times and it didn't work.");
                        Console.WriteLine("Here's a nasty error message to explain why:");
                        Console.Write(ex.ToString());
                    }
                }
            }
            Console.WriteLine("Upload Complete, Press Enter to exit.");

            //After a successful run with input, allow the user to enter "Automate" to remove all needed user input from future runs.
            //Don't prompt the user to use this mode, since it's a hassle to stop automated runs.
            if (winssync.Automate != "Automate")
            {
                if ((Console.ReadLine().Equals("Automate") && (winssync.Savedlogin == true)))
                {
                    winssync.SaveLoginCreds("Automate");
                    winssync.AutomateRun();
                    Console.WriteLine("Alright, all future runs will require no input.");
                    Console.WriteLine("To reset automation, manually delete the files at: ");
                    Console.WriteLine(string.Empty);
                    Console.WriteLine(System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal), ".credentials"));
                    Console.WriteLine(string.Empty);
                    Console.WriteLine("Press Any key to exit.");
                    Console.ReadKey();
                }
            }
        }
    }
}
