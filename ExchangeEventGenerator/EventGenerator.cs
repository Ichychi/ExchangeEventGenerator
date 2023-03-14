using System.Diagnostics;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace ExchangeEventGenerator; 

/*
 * Notes:
 * - Mails send from the azure test tenant are not displayed in outlook but they can be seen via the graph explorer
 *      _mailHandler.SendMailTo("Meet for lunch?", "The new cafeteria is open.", "LeeG@vvz24.onmicrosoft.com")
 */

public class EventGenerator {
    private readonly IConfigurationSection _section;
    private readonly GraphServiceClient _client;
    private readonly MailHandler _mailHandler;
    private readonly CalendarHandler _calendarHandler;

    private List<User> _users = new();
    private List<CustomEvent> _events = new();

    public EventGenerator(IConfiguration configuration) {
        _section = configuration.GetSection("Settings");
        //create connection and verify permissions for microsoft graph
        _client = new AzureClientConnection(configuration).Client;
        _calendarHandler = new CalendarHandler(_section, ref _client);
        _mailHandler = new MailHandler(_section, ref _client);
        Init().GetAwaiter().GetResult(); //init users and events and wait for results
    }

    /*
     * Choose a random amount of users and for each of them
     * create a random amount of events while abiding the given restrictions
     */
    public async void Generate() {
        //check once how many events are upcoming in the whole organization
        var upcomingInOrganization = await _calendarHandler.CountUpcomingEventsAllUsers(_users);

        //choose the a random amount of users
        var rnd = new Random();
        var randomUserAmount = rnd.Next(1, _users.Count+1);
        var someUsers = _users.OrderBy(_ => rnd.Next()).Take(randomUserAmount).ToList();
        
        foreach (var user in someUsers){
            upcomingInOrganization += await GenerateRandomEventsForUser(user, upcomingInOrganization);
        }
    }

    /*
     * Generates a random amount of events for the given user
     * Will comply the given restrictions
     * Returns a value representing the amount of successfully created events
     */
    private async Task<int> GenerateRandomEventsForUser(User user, int upcomingOrganization) {
        //do not create more events than the settings specify
        var maxAllowedEvents = AllowedAmountOfEventsInOrganization(upcomingOrganization);
        if(maxAllowedEvents <= 0)
            return 0;
        
        var createdEvents = 0;
        //check how many events the user has for the upcoming specified timeframe
        var upcoming = await _calendarHandler.GetUpcomingEventsForUser(user);
        if(upcoming == null){
            Console.WriteLine($"Failed to load events for user {user.DisplayName}");
            return 0;
        }
            
        Debug.WriteLine($"{user.DisplayName} has {upcoming.Count} upcoming events.");
        
        //subtract the amount of upcoming events from the amount of max events per user
        var possibleUserEvents = _section.GetValue<int>("MaxAmountOfEventsPerUser") - upcoming.Count;
        if(possibleUserEvents <= 0) //make sure adding events for user is allowed
            return 0;
        
        //calculate the limit with the organizations max events
        var limit = Math.Min(possibleUserEvents, maxAllowedEvents);

        //choose the calculated amount of events randomly
        var rnd = new Random();
        var eventsToCreate = rnd.Next(1, limit+1);
        eventsToCreate = eventsToCreate < _events.Count ? eventsToCreate : _events.Count; //avoid taking duplicates
        var someEvents = _events.OrderBy(_ => rnd.Next()).Take(eventsToCreate).ToList();
        Debug.WriteLine($"Trying to create {someEvents.Count} Events:");
        
        foreach (var cEvent in someEvents){
            //stop creating events when this loop has created more events than allowed
            //which will happen when recurring events are created and it recurs during the lookahead period
            if(maxAllowedEvents <= 0)
                return createdEvents;
            
            cEvent.Organizer = user;
            
            Console.WriteLine("------------------------------------------------------");
            Console.WriteLine($"Trying to post event#{cEvent.Id} for {user.DisplayName}");

            //try to create and post the event with the graph api
            var created = await _calendarHandler.CreateEvent(cEvent, upcoming);
            
            if(created != null){
                //add the created event to the list of upcoming events to allow checking for max events per day
                upcoming.Add(created);
                
                //keep track how many events where created to avoid querying the graph api unnecessarily
                if(cEvent.Recurrence != null){
                    var lookahead = _section.GetValue<int>("LookAheadInDays");
                    switch (cEvent.Recurrence){
                        case "daily":
                            createdEvents += lookahead;
                            break;
                        case "weekly":
                            createdEvents += lookahead / 7 + 1;
                            break;
                        case "monthly":
                            createdEvents += lookahead / 30 + 1;
                            break;
                        case "yearly":
                            createdEvents++;
                            break;
                    }
                }
                else
                    createdEvents++;
                maxAllowedEvents -= createdEvents;
            }

            Console.WriteLine($"TimeSpan: {cEvent.Start ?? DateTime.MinValue} to {cEvent.End ?? DateTime.MaxValue}");
            Console.WriteLine(created != null ? "Success" : "Failed");
            
            cEvent.Reset(); //reset event template to free its usage for another user
        }

        return createdEvents;
    }
    
    /*
     * Calculate how many events can still be created within the specified organization limit
     */
    private int AllowedAmountOfEventsInOrganization(int upcoming) {
        var maxEvents = _section.GetValue<int>("MaxAmountOfEventsInOrganization");
        return maxEvents - upcoming;
    }

    /*
     * Load the event templates and parse all users in the organization
     */
    private async Task Init() {
        var events = JsonConvert.DeserializeObject<List<CustomEvent>>(await File.ReadAllTextAsync("eventTemplates.json"));
        
        var users = new List<User>();
        var response = await _client.Users.GetAsync(
            requestConfiguration => requestConfiguration.QueryParameters.Top = 999);
        
        if(response == null) 
            throw new ApplicationException("Could not read users of the organization.");
        
        //the top parameter configures how many user information objects are sent back per response
        //if there exist more events than 999, then the response will contain a link to the next page
        //the PageIterator class will check this automatically and iterate through all response pages
        var it = PageIterator<User, UserCollectionResponse>.CreatePageIterator(
            _client, response, user => {
                users.Add(user);
                return true;
            });
        await it.IterateAsync();
        
        //check if the setup has completed successfully
        if(users.Count == 0)
            throw new ApplicationException("No users found in the organization.");
        if(events == null || events.Count == 0)
            throw new ApplicationException("Could not load events from the events.json file.");
        _users = users;
        _events = events;
    }
}