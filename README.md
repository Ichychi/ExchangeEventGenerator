# An event generator for Microsoft Exchange
[![Hack Together: Microsoft Graph and .NET](https://img.shields.io/badge/Microsoft%20-Hack--Together-orange?style=for-the-badge&logo=microsoft)](https://github.com/microsoft/hack-together)

This is a service worker which can create "random" events for every member of an organization utilizing the Microsoft Graph API.

Basically simulating activity within an organization.

## How it works
The service worker will run every [X minutes](ExchangeEventGenerator/appsettings.json) to create events:
- Events are based on event templates [eventTemplates.json](ExchangeEventGenerator/eventTemplates.json)
- A random date and time within the [lookahead](ExchangeEventGenerator/appsettings.json) is applied to each of these events
- A random amount of users within the organization is chosen
- A random amount of events is created for them

## Requirements
- .NET v7.0
- Azure subscription with admin privileges for the target tenant

## Setup
1. Add a new app registration to your target tenant
2. Create a certificate on your local machine and upload it for the authentication with azure  
e.g. follow the first half of this [guide](https://blogs.aaddevsup.xyz/2020/07/using-msal-net-to-perform-the-client-credentials-flow-with-a-certificate-instead-of-a-client-secret-in-a-netcore-console-appliction/)
3. Add the following app permissions on azure: "User.Read.All" and "Calendars.ReadWrite"
4. Add your tenant id, client id and certificate thumbprint to the [appsettings file](ExchangeEventGenerator/appsettings.json)
5. In that file you can also adjust the "settings" section to your liking
6. *(Optional)* Adjust [event templates](ExchangeEventGenerator/eventTemplates.json):

### Event Template Layout
|       Property        |                                    Description                                    |                                                                                     Expected Value                                                                                     |
|:---------------------:|:---------------------------------------------------------------------------------:|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------:|
|          Id           |                       An identifier for the event template                        |                                                                             A number unique to this event                                                                              |
|        Subject        |                             The subject of the event                              |                                                                             A string defining the subject                                                                              |
|        Content        |                             The content of the event                              |                                                                             A string defining the content                                                                              |
|       IsAllDay        |                         Defining if the event is all day                          |                                                                                  Either true or false                                                                                  |
|       Duration        |             Duration of the event, only required if IsAllDay is false             |                                                                A string defining the duration (e.g. "2:00" for 2 hours)                                                                |
|       Reminder        |              How many minutes ahead should the event remind the user              |                                                                             A number defining the minutes                                                                              |
|      Recurrence       |                         The recurrence type of the event                          |                                                                   "" or "daily" or "weekly" or "monthly" or "yearly"                                                                   |
|      DaysOfWeek       | The day of the week the event should start at, only required if recurrence is set |                                                  A string containing weekdays separated by commas (e.g. "monday" or "monday,tuesday")                                                  |
|  NumberOfOccurrences  |    How many times the event should recurr, only required if recurrence is set     |                                                              A number defining how many times to repeat a recurring event                                                              |
|      Importance       |                            The importance of the event                            |                                                                              "low" or "normal" or "high"                                                                               |
|        ShowAs         |                       The status of the event for its owner                       |                                                      "free" or "tentative" or "busy" or "oof" or "workingElsewhere" or "unknown"                                                       |
|      Attachments      |                               A list of attachments                               | A string containing attachment names in the [Attachment folder](ExchangeEventGenerator/Attachments) separated by commas (e.g. "Lorem Ipsum.pdf" or "Lorem Ipsum.pdf,Lorem Picsum.jpg") |

## Usage
```
dotnet run
```

## Notes
- If it states "Failed" for the creation of an event, then that means that it could not find a timeslot within the specified [restrictions](ExchangeEventGenerator/appsettings.json).
- Try to avoid adding many reccuring events *(especially daily ones)* as it will fill the calender boringly.
- Adding attendees to an event is currently not implemented. I was not able to test this feature, as i was using the azure testversion, where sending mails via the graph API is not allowed.
