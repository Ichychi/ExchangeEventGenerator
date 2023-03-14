using System.Diagnostics;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Newtonsoft.Json;

namespace ExchangeEventGenerator;

public class CustomEvent {
    private static readonly int[] ValidMinutes = {0, 15, 30, 45};
    [JsonProperty]
    public int Id { get; set; } //not the actual exchange event id
    [JsonProperty]
    public string? Subject { get; set; }
    [JsonProperty]
    public string? Content { get; set;}
    [JsonProperty]
    public bool IsAllDay { get; set;}
    [JsonProperty]
    public string? Duration { get; set;}
    [JsonProperty]
    public int Reminder { get; set; }
    [JsonProperty]
    public string? Recurrence { get; set;}
    [JsonProperty]
    public string? DaysOfWeek { get; set;}
    [JsonProperty]
    public int NumberOfOccurrences { get; set;}
    [JsonProperty]
    public string? Importance { get; set;}
    [JsonProperty]
    public string? ShowAs { get; set;}
    [JsonProperty]
    public string? Attachments { get; set; }
    
    public User? Organizer { get; set; }
    private List<Attendee> Attendees => new List<Attendee>(); //TODO change this once setting attendees works (sending mails?)
    public DateTime? Start { get; set; }
    public DateTime? End { get; set; }
    private TimeZoneInfo ZoneInfo => TimeZoneInfo.Local; //TODO? add support for different timezones

    /*
     * Resetting a custom event to enable the creation at different times for different users
     * without the need of recreating the entire object
     */
    public void Reset() {
        Organizer = null;
        //Attendees = new List<Attendee>();
        Start = null;
        End = null;
        //TimeZoneInfo = null;
    }

    /*
     * Try to create the PatternedRecurrence object of a Graph API event based on this custom object
     */
    private PatternedRecurrence? GetPatternedRecurrence() {
        Debug.WriteLineIf(Start == null, "Start time is null.");
        //if recurrence is not specified, returning null as required by graph api
        return string.IsNullOrWhiteSpace(Recurrence) || Start == null ? 
            null :
            new PatternedRecurrence {
            Pattern = new RecurrencePattern {
                Type = GetPatternType(),
                DayOfMonth = GetDayOfMonth(),
                Month = GetMonth(),
                DaysOfWeek = GetDaysOfWeek(),
                FirstDayOfWeek = DayOfWeekObject.Monday,
                Interval = 1 //currently not supporting every x days, biweekly, quarterly, every x years
            },
            Range = new RecurrenceRange {
                Type = RecurrenceRangeType.Numbered,
                NumberOfOccurrences = NumberOfOccurrences,
                StartDate = new Date(Start.Value)
            }
        };
    }

    /*
     * Parse the Recurrence Property and return the RecurrencePatternType
     */
    private RecurrencePatternType? GetPatternType() {
        var lower = Recurrence?.ToLower() ?? "";
        return lower switch {
            "daily" => RecurrencePatternType.Daily,
            "weekly" => RecurrencePatternType.Weekly,
            "monthly" => RecurrencePatternType.AbsoluteMonthly,
            "yearly" => RecurrencePatternType.AbsoluteYearly,
            _ => null
        };
    }

    private int GetDayOfMonth() {
        return Start?.Day ?? 0;
    }

    private int GetMonth() {
        return Start?.Month ?? 0;
    }
    
    /*
     * Construct the ItemBody of the event
     */
    private ItemBody GetItemBody() {
        return new ItemBody {
            Content = Content,
            ContentType = BodyType.Text
        };
    }

    /*
     * Return a DateTimeTimeZone representation of the Start Property
     */
    private DateTimeTimeZone GetStartTime() {
        Debug.WriteLineIf(Start == null || ZoneInfo == null, "Start time is null.");
        return new DateTimeTimeZone {
            DateTime = $"{Start:yyyy-MM-dd}T{Start:HH:mm:sss}",
            TimeZone = ZoneInfo?.Id ?? TimeZoneInfo.Local.Id
        };
    }
    
    /*
     * Return a DateTimeTimeZone representation of the End Property
     */
    private DateTimeTimeZone GetEndTime() {
        Debug.WriteLineIf(End == null || ZoneInfo == null, "End time is null.");
        return new DateTimeTimeZone {
            DateTime = $"{End:yyyy-MM-dd}T{End:HH:mm:sss}",
            TimeZone = ZoneInfo?.Id ?? TimeZoneInfo.Local.Id
        };
    }

    /*
     * Parse the DaysOfWeek property and return a list of graph api DayOfWeekObjects
     */
    private List<DayOfWeekObject?> GetDaysOfWeek() {
        var daysOfWeek = new List<DayOfWeekObject?>();
        var lower = DaysOfWeek?.ToLower() ?? "";
        if(string.IsNullOrWhiteSpace(lower))
            return daysOfWeek;
        
        //expected format: e.g. "monday,thursday,friday"
        var days = lower.Split(',').ToList();
        if(days.Contains("saturday") || days.Contains("sunday")){
            throw new InvalidDataException("Can not create events on saturday or sunday.");
        }

        foreach (var day in days){
            switch (day){
                case "monday":
                    daysOfWeek.Add(DayOfWeekObject.Monday);
                    break;
                case "tuesday":
                    daysOfWeek.Add(DayOfWeekObject.Tuesday);
                    break;
                case "wednesday":
                    daysOfWeek.Add(DayOfWeekObject.Wednesday);
                    break;
                case "thursday":
                    daysOfWeek.Add(DayOfWeekObject.Thursday);
                    break;
                case "friday":
                    daysOfWeek.Add(DayOfWeekObject.Friday);
                    break;
                default:
                    throw new InvalidDataException($"Invalid day specified: {day}");
            }
        }

        return daysOfWeek;
    }

    /*
     * Create a Recipient graph api object for the organizer of this event
     */
    private Recipient? GetOrganizer() {
        Debug.WriteLineIf(Organizer == null, "Organizer is null.");
        return Organizer == null
            ? null
            : new Recipient {
                EmailAddress = new EmailAddress {
                    Address = Organizer.Mail,
                    Name = Organizer.DisplayName
                }
            };
    }

    private Importance? GetImportance() {
        var success = Enum.TryParse(Importance, true, out Importance parsed);
        return success ? parsed : null;
    }

    private FreeBusyStatus? GetShowAs() {
        var success = Enum.TryParse(ShowAs, true, out FreeBusyStatus parsed);
        return success ? parsed : null;
    }
    
    /*
     * Calculates the start and end time of this custom event randomly,
     * while respecting the following Properties:
     * - IsAllDay
     * - Duration
     * - Recurrence
     * - DaysOfWeek
     */
    public void CalculateStartAndEnd(IConfigurationSection section) {
        //start by calculating the day of the event
        DateTime? day;
        if(!string.IsNullOrWhiteSpace(Recurrence)){
            if(Recurrence == "daily")
                day = DateTime.Today;
            else{
                var days = GetDaysOfWeek();
                if(days.Count < 1)
                    throw new InvalidDataException("Recurrence was set but DaysOfWeek does not contain valid days.");
                var firstDay = days.Min();
                if(firstDay == null)
                    throw new InvalidDataException("Could not find valid day to set as first day of the event.");
                day = GetNextWeekday(DateTime.Today, firstDay.Value);
            }
        }else{
            day = GetRandomDay(section);
        }

        if(day == null)
            throw new TaskCanceledException("Could not find valid day for event.");
        
        //then calculate the time of the event
        var time = GetRandomTime(day.Value);
        if(time == null)
            throw new TaskCanceledException("Could not find valid time for event.");

        Start = time.Value;

        //finally calculate the end date of the event
        if(IsAllDay){
            End = Start.Value.AddDays(1);
        }else{
            var duration = TimeSpan.Parse(Duration ?? "0");
            End = Start.Value.AddHours(duration.Hours).AddMinutes(duration.Minutes); 
        }
    }
    
    //given a start date and a target day, it will calculate the date of upcoming next day
    private static DateTime GetNextWeekday(DateTime start, DayOfWeekObject day) {
        return start.AddDays(((int) day - (int) start.DayOfWeek + 7) % 7);
    }

    /*
     * Trys to find a random day during the upcoming days, excluding saturday and sunday
     */
    private static DateTime? GetRandomDay(IConfigurationSection section) {
        //return a random date depending on the lookahead
        var rnd = new Random();
        var lookAhead = section.GetValue<int>("LookAheadInDays");
        DateTime? rndDay = null;
        for (var i = 0; i < 10; i++){
            var newDate = DateTime.Today.AddDays(rnd.Next(0, lookAhead + 1));
            if(IsValidDay(newDate)){
                rndDay = newDate;
                break;
            }
        }
        return rndDay;
    }

    /*
     * Trys to find a random time on the specified day
     * Valid hours are from 7:00-18:00
     * Valid minutes are defined in the ValidMinutes array
     */
    private DateTime? GetRandomTime(DateTime day) {
        if(IsAllDay) //IsAllDay = true => "start and end time must be set to midnight and be in the same time zone"
            return day;
        
        //get a random minute definition from the ValidMinutes array
        var rnd = new Random();
        var minute = ValidMinutes[rnd.Next(0, ValidMinutes.Length)];
        
        if(!TimeSpan.TryParse(Duration, out var duration))
            throw new InvalidDataException($"Could not parse Duration for event#{Id}");

        //get a random hour between 7 and 18, but make sure the event fits in
        var hour = rnd.Next(7, 19 - (duration.Hours == 0 ? 1 : duration.Hours));
        var randomTime = day.AddHours(hour).AddMinutes(minute);

        return IsValidStartTime(randomTime) ? randomTime : null;
    }

    /*
     * Trys to create a Graph API event object for this custom event object
     */
    public Event? ToEvent(List<Event> upcoming, int maxEventsPerDay) {
        //error checking first
        if(!IsValidEvent(upcoming, maxEventsPerDay))
            return null;

        return new Event {
            Subject = Subject,
            Body = GetItemBody(),
            IsAllDay = IsAllDay,
            Start = GetStartTime(),
            End = GetEndTime(),
            Recurrence = GetPatternedRecurrence(),
            ReminderMinutesBeforeStart = Reminder,
            Attendees = Attendees,
            Organizer = GetOrganizer(),
            Importance = GetImportance(),
            ShowAs = GetShowAs()
        };
    }
    
    /*
     * A custom event is valid when every user specific property has been set
     * And the custom event fits regarding the specified max amount of events per day
     */
    private bool IsValidEvent(List<Event> upcoming, int maxEventsPerDay) {
        if(Organizer == null || Start == null || End == null || ZoneInfo == null)
            return false;
        
        //check the day of this event and search for events on same day
        var upcomingEvents = upcoming.Count(e => e.Start.ToDateTime().Date == Start.Value.Date);
        return upcomingEvents < maxEventsPerDay;
    }

    private static bool IsValidDay(DateTime dateTime) {
        return dateTime.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday);
    }

    private static bool IsValidStartTime(DateTime dateTime) {
        //TODO? add event conflicting checks
        return true;
    }

    public bool HasAttachment() {
        return !string.IsNullOrWhiteSpace(Attachments);
    }

    public List<Attachment> GetAttachments() {
        var attachments = new List<Attachment>();
        var lower = Attachments?.ToLower() ?? "";
        if(string.IsNullOrWhiteSpace(lower))
            return attachments;
        
        //expected format: e.g. "Attachment1.txt,Attachment2.pdf,Attachment3.jpg"
        const string folder = "Attachments/";
        var files = lower.Split(',').ToList();

        foreach (var file in files){
            var base64 = Convert.ToBase64String(File.ReadAllBytes(folder + file));
            attachments.Add(new Attachment {
                OdataType = "#microsoft.graph.fileAttachment",
                Name = file,
                AdditionalData = new Dictionary<string, object> { {
                        "contentBytes" , base64
                    }
                }
            });
        }
        return attachments;
    }
}