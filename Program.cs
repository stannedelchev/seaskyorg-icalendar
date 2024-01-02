using System.CommandLine;
using System.Globalization;
using System.Text.RegularExpressions;
using AngleSharp;
using AngleSharp.Dom;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;

var urlOption = new Option<Uri>("--url",
                                 () => Url.Create("http://www.seasky.org/astronomy/astronomy-calendar-2024.html"),
                                 "The http://seasky.org page URL to convert to .ical format.");
var rootCommand = new RootCommand("Convert an SeaSky.org page with astronomical events to .ical format for usage with various calendar software.")
{
    urlOption
};
rootCommand.SetHandler(BuildCalendar, urlOption);
return await rootCommand.InvokeAsync(args);

async Task BuildCalendar(Uri url)
{
    var angleSharpDocument = await BrowsingContext.New(Configuration.Default.WithDefaultLoader())
                                                  .OpenAsync(url.ToString());
    var year = int.Parse(angleSharpDocument.QuerySelector("h1")!.TextContent[^4..]); // "Astronomy Calendar of Celestial Events for Calendar Year 2024"

    var events = angleSharpDocument.QuerySelectorAll("div#right-column-content li p")
                                   .Select(e =>
                                   {
                                       var summary = GetSummary(e);
                                       var description = GetDescription(e);
                                       var (start, end) = GetDate(e, year);
                                       return new Event(summary, description, start, end);
                                   });

    var calendar = new Ical.Net.Calendar();
    calendar.Events.AddRange(events.Select(e => new CalendarEvent
    {
        Start = new CalDateTime(e.StartDate),
        End = new CalDateTime(e.EndDate),
        Uid = Guid.NewGuid().ToString("D"),
        Summary = e.Summary,
        Description = e.Description
    }));

    var serializedCalendar = new CalendarSerializer().SerializeToString(calendar);
    Console.WriteLine(serializedCalendar);
}

string GetSummary(IElement element)
{
    return element.QuerySelector("span.title-text")!
                  .TextContent
                  .Trim('.')
                  .Trim();
}

string GetDescription(IElement element)
{
    var content = element.InnerHtml;
    var descriptionStartIndex = content.LastIndexOf("</span>", StringComparison.InvariantCulture) + "</span>".Length;
    return content[descriptionStartIndex..].Trim();
}

(DateTime Start, DateTime End) GetDate(IElement element, int year)
{
    var content = element.QuerySelector("span.date-text")!.TextContent.Trim('.').Trim();
    var dateRange = Regex.Match(content, @"(?<month>[a-zA-Z]+) (?<dayStart>\d{1,2})(, )?(?<dayEnd>\d{1,2})?");

    if (!dateRange.Success ||
        !dateRange.Groups["month"].Success ||
        !dateRange.Groups["dayStart"].Success)
    {
        throw new InvalidOperationException();
    }

    var month = dateRange.Groups["month"];
    var dayStart = dateRange.Groups["dayStart"];
    var dayEnd = dateRange.Groups["dayEnd"];

    var startDate = DateTime.ParseExact($"{month} {dayStart} {year}", "MMMM d yyyy", CultureInfo.InvariantCulture);
    var endDate = dateRange.Groups["dayEnd"].Success
        ? DateTime.ParseExact($"{month} {dayEnd} {year}", "MMMM d yyyy", CultureInfo.InvariantCulture)
        : startDate;

    return (startDate, endDate.AddDays(1));
}

internal record Event(string Summary, string Description, DateTime StartDate, DateTime EndDate);