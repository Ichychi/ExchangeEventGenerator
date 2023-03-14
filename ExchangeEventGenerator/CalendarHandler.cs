using System.Diagnostics;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

/* Note
 * Adding attendees to an event does not add the event for the attendee
 * This might be because of the test tenant which does not allow sending mails
 */

namespace ExchangeEventGenerator; 

public class CalendarHandler {
    private readonly IConfigurationSection _section;
    private readonly GraphServiceClient _client;

    public CalendarHandler(IConfigurationSection section, ref GraphServiceClient client) {
        _section = section;
        _client = client;
    }
    
    /*
     * Calculate the start/end of the given CustomEvent then convert it to a Graph API event and post it
     * Also makes sure the specified limit of max events per day is taken into consideration
     * Returns the result of the post request
     */
    public async Task<Event?> CreateEvent(CustomEvent cEvent, List<Event> upcoming) {
        try{
            cEvent.CalculateStartAndEnd(_section);
        }catch (TaskCanceledException){//TODO? create own exception
            Console.WriteLine("Could not calculate start and end time for event.");
            //this case happens when it was not possible to find a free timeslot for the event on the users calendar
            return null; 
        }

        if(cEvent.Organizer == null){
            Console.WriteLine("Events without an organizer are invalid.");
            return null;
        }

        //make sure the event is valid
        var request = cEvent.ToEvent(upcoming, MaxEventsPerDay());
        if(request == null)
            return null;
        var response = await _client.Users[cEvent.Organizer.Id].Calendar.Events.PostAsync(request);
        if(response == null)
            return null;

        if(!cEvent.HasAttachment())
            return response;
        foreach (var attachment in cEvent.GetAttachments()){
            await AddAttachment(cEvent.Organizer, response, attachment);
        }
        return response;
    }

    private async Task<Attachment?> AddAttachment(User user, Event response, Attachment attachment) {
        return await _client.Users[user.Id].Events[response.Id].Attachments.PostAsync(attachment);
    }

    /*
     * Deletes all events within the specified lookahead range for the given user
     */
    public async Task DeleteAllUpcomingEvents(User user) {
        //get all upcoming events for user
        var events = await GetUpcomingEventsForUser(user);
        if(events != null)
            //delete them by id
            await Task.WhenAll(events.Select(async userEvent =>
                await _client.Users[user.Id].Calendar.Events[userEvent.Id].DeleteAsync()));
    }
    
    /*
     * Deletes all events for the given user
     */
    public async Task DeleteAllEvents(User user) {
        Console.WriteLine($"Deleting all events for {user.DisplayName}");
        //get all events for user
        var events = await _client.Users[user.Id].Calendar.Events.GetAsync();
        //delete them by id
        await Task.WhenAll(events?.Value?.Select(async userEvent => 
            await _client.Users[user.Id].Calendar.Events[userEvent.Id].DeleteAsync()) ?? Array.Empty<Task>());
    }

    private int MaxEventsPerDay() {
        return _section.GetValue<int>("MaxEventsOnSameDay");
    }
    
    /*
     * Returns a list of all upcoming(=lookahead setting) events for the given user
     * As the get request is done on the calendar view, this also sees recurring events individually
     */
    public async Task<List<Event>?> GetUpcomingEventsForUser(User user) {
        var events = new List<Event>();
        var lookAhead = _section.GetValue<int>("LookAheadInDays");
        var response = await _client.Users[user.Id].CalendarView.GetAsync(
            requestConfiguration => {
            requestConfiguration.QueryParameters.StartDateTime = DateTime.Today.ToString("yyyy-MM-ddTHH:mm:ss");
            requestConfiguration.QueryParameters.EndDateTime = DateTime.Today.AddDays(lookAhead+1).ToString("yyyy-MM-ddTHH:mm:ss");
            requestConfiguration.QueryParameters.Top = 999;
            });

        if(response == null) 
            return null;
        
        //the top parameter configures how many events are sent back per response, here set to 999 (max of graph api)
        //if there exist more events than 999, then the response will contain a link to the next page
        //the PageIterator class will check this automatically and iterate through all response pages
        var it = PageIterator<Event, EventCollectionResponse>.CreatePageIterator(
            _client, response, ev => {
                events.Add(ev);
                return true;
            });
        await it.IterateAsync();
        return events;
    }
    
    /*
     * Returns a list of all upcoming(=lookahead setting) events for all users
     * As the get request is done on the calendar view, this also sees recurring events individually
     */
    public async Task<int> CountUpcomingEventsAllUsers(List<User> users) {
        //make use of a batch request to speed things up
        var batchRequestContent = new BatchRequestContent(_client);
        //batch request requires setting up request ids to enable querying the response
        var requestIds = new List<string>(); 
        foreach (var user in users){
            requestIds.Add(await batchRequestContent.AddBatchRequestStepAsync(UpcomingEventsRequestForUser(user)));
        }
        var result = await _client.Batch.PostAsync(batchRequestContent);
        
        var upcomingEvents = 0;
        foreach (var requestId in requestIds){
            //get each response based on the request id and then use an iterator to include pagination responses
            var response = await result.GetResponseByIdAsync<EventCollectionResponse>(requestId);
            var it = PageIterator<Event, EventCollectionResponse>.CreatePageIterator(
                _client, response, _ => {
                    upcomingEvents++;
                    return true;
                });
            await it.IterateAsync();
        }
        Debug.WriteLine($"Amount of upcoming events: {upcomingEvents}");
        return upcomingEvents;
    }

    /*
     * Helper function to set up the batch request above
     * Constructs the RequestInformation based on the lookahead
     */
    private RequestInformation UpcomingEventsRequestForUser(User user) {
        var lookAhead = _section.GetValue<int>("LookAheadInDays");
        return _client.Users[user.Id].CalendarView.ToGetRequestInformation(
            requestConfiguration => {
                requestConfiguration.QueryParameters.StartDateTime = DateTime.Today.ToString("yyyy-MM-ddTHH:mm:ss");
                requestConfiguration.QueryParameters.EndDateTime = DateTime.Today.AddDays(lookAhead+1).ToString("yyyy-MM-ddTHH:mm:ss");
                requestConfiguration.QueryParameters.Top = 999;
            });
    }
}