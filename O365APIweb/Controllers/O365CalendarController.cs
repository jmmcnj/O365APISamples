using Microsoft.Office365.Exchange;
using Microsoft.Office365.OAuth;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace O365APIweb.Controllers
{

    //local class to handle cal event data
    public class CalendarEvent
    {
        public string Subject { get; set; }
        public string Location { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
    }

    public class O365CalendarController : Controller
    {
        //private members
        private const string ServiceResourceId = "https://outlook.office365.com";
        private static readonly Uri ServiceEndpointUri = new Uri("https://outlook.office365.com/ews/odata");
        private static DiscoveryContext _discoveryContext;


        public async Task<ExchangeClient> GetExchangeClient()
        {
            // Create the discovery context if it doesn't already exist.
            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            // Authenticate and retrieve the tenant ID and user ID.
            var discoverResult = await _discoveryContext.DiscoverResourceAsync(ServiceResourceId);

            string refreshToken = new SessionCache().Read("RefreshToken");

            Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential creds =
                new Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential(
                    _discoveryContext.AppIdentity.ClientId, _discoveryContext.AppIdentity.ClientSecret);

            return new ExchangeClient(ServiceEndpointUri, async () =>
            {
                // Get the access token based on the refresh token.
                return (await _discoveryContext.AuthenticationContext.AcquireTokenByRefreshTokenAsync(
                    refreshToken, creds, ServiceResourceId)).AccessToken;
            });
        }

        // GET: O365Calendar
        public async Task<ActionResult> Index()
        {
            List<CalendarEvent> myEvents = new List<CalendarEvent>();

            try
            {
                // Call the GetExchangeClient method, which will authenticate
                // the user and create the ExchangeClient object.
                var client = await GetExchangeClient();
                if (client == null)
                {
                    return View(myEvents);
                }

                // Use the ExchangeClient object to call the Calendar API.
                // Get all events that have an end time after now.
                var eventsResults = await (from i in client.Me.Events
                                           where i.End >= DateTimeOffset.UtcNow
                                           select i).Take(10).ExecuteAsync();

                // Order the results by start time.
                var events = eventsResults.CurrentPage.OrderBy(e => e.Start);

                // Create a CalendarEvent object for each event returned
                // by the API.
                foreach (Event calendarEvent in events)
                {
                    CalendarEvent newEvent = new CalendarEvent();
                    newEvent.Subject = calendarEvent.Subject;
                    newEvent.Location = calendarEvent.Location.DisplayName;
                    newEvent.Start = calendarEvent.Start.GetValueOrDefault().DateTime;
                    newEvent.End = calendarEvent.End.GetValueOrDefault().DateTime;

                    myEvents.Add(newEvent);
                }
            }
            // Required exception handling to make redirection work.
            catch (RedirectRequiredException redir)
            {
                return Redirect(redir.RedirectUri.ToString());
            }

            return View(myEvents);
        }

    }
}