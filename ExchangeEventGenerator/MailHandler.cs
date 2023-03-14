using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace ExchangeEventGenerator; 

public class MailHandler {
    private readonly GraphServiceClient _client;

    public MailHandler(IConfigurationSection section, ref GraphServiceClient client) {
        _client = client;
    }
    
    /*
     * Simple mail sending function, needs to be tested on a different tenant where sending mails is possible
     */
    public async void SendMailTo(string subject, string content, string recipient) {
        var requestBody = new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody {
            Message = new Message {
                Subject = subject,
                Body = new ItemBody {
                    ContentType = BodyType.Text,
                    Content = content
                },
                ToRecipients = new List<Recipient> {
                    new Recipient {
                        EmailAddress = new EmailAddress {
                            Address = recipient
                        }
                    }
                }
            }
        };
        await _client.Me.SendMail.PostAsync(requestBody);
    }
}