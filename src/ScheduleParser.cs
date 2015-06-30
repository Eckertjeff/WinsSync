 using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

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
        public string StartTime; // todo parse into Time struct
        public string EndTime; // todo parse into Time struct
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

        static void Main(string[] args)
        {
            var originalColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Cyan;

            var file = File.ReadAllText("Welcome.html");
            var endTable = "</table>";
            var startTable = "<table";

            int index = 0;

            Console.WriteLine("Parsing DOM for HTML Tables");
            // Parse the DOM, find our tables
            while ((index = file.IndexOf(startTable, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                var tableContentEndIndex = file.IndexOf(endTable, index);
                var tableContent = file.Substring(index, tableContentEndIndex - index + endTable.Length);

                Console.WriteLine("Found a table DOM element!");

                m_Tables.Add(tableContent);
            }

            // Identify the table that's actually relevant to us
            var keyword = "Schedule Times:";
            var oneWeWantList = m_Tables.Where(s => s.Contains(keyword));

            if (oneWeWantList.Count() > 1)
            {
                throw new Exception("Our hacky ass code found more than one 'right' one. Do it better.");
            }

            var correctTable = oneWeWantList.FirstOrDefault();

            if (correctTable == null)
            {
                throw new Exception("Our shitty ass code didn't find the right one at all. #NailedIt.");
            }

            Console.WriteLine("Identified which Table is the Schedule...");

            // Now we gotta parse our table for the values we want
            const string rowStart = "<tr";
            const string cellStart = "<td";
            const string rowEnd = "</tr>";
            const string cellEnd = "</td>";

            var rows = new List<Row>();

            index = 0;
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

            Console.WriteLine("Parse complete... All your HTML belongs to us.");
            Console.WriteLine(string.Empty);
            Console.ForegroundColor = originalColor;

            // Now that we have all the data in a parsable format
            // w need to parse the "rows" object
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

                schedule.Add(workDay);
                dayIndex++;
            }

            // Display our results
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

            Console.ReadKey();
            }
        }
    }
