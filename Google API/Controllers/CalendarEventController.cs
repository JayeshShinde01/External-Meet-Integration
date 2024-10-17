using GoogleAPI.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using RestSharp;
using System;
using System.Web.Mvc;

namespace GoogleAPI.Controllers
{
    public class CalendarEventController : Controller
    {
        [HttpPost]
        public ActionResult CreateEvent(Event calendarEvent, string externalMeetingLink)
        {
            var tokenFile = "C:\\Users\\DELL\\source\\repos\\google-calendar-integration\\Google API\\Files\\tokens.json";
            var tokens = JObject.Parse(System.IO.File.ReadAllText(tokenFile));

            RestClient restClient = new RestClient("https://www.googleapis.com/calendar/v3/calendars/primary/events");

            RestRequest request = new RestRequest(Method.POST);

            // Parse and format DateTime strings to ISO 8601 format
            calendarEvent.Start.DateTime = DateTime.ParseExact(calendarEvent.Start.DateTime, "yyyy-MM-ddTHH:mm", System.Globalization.CultureInfo.InvariantCulture).ToString("yyyy-MM-dd'T'HH:mm:ss.fffK");
            calendarEvent.End.DateTime = DateTime.ParseExact(calendarEvent.End.DateTime, "yyyy-MM-ddTHH:mm", System.Globalization.CultureInfo.InvariantCulture).ToString("yyyy-MM-dd'T'HH:mm:ss.fffK");

            // Add external meeting link to the location field
            calendarEvent.Location = externalMeetingLink;

            // Serialize the event object into JSON
            var model = JsonConvert.SerializeObject(calendarEvent, new JsonSerializerSettings
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver()
            });

            // Add necessary headers and parameters to the request
            request.AddQueryParameter("key", "Your API Key"); // Use your API key
            request.AddHeader("Authorization", "Bearer " + tokens["access_token"]); // Use the access token
            request.AddHeader("Accept", "application/json");
            request.AddHeader("Content-Type", "application/json");

            // Attach the serialized event as the body of the request
            request.AddParameter("application/json", model, ParameterType.RequestBody);

            // Execute the request
            var response = restClient.Execute(request);

            // Check if the event creation was successful
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                var createdEvent = JObject.Parse(response.Content);

                // Extract event details if needed
                var eventId = createdEvent["id"]?.ToString();

                // Redirect to a success page or display the link
                return RedirectToAction("Index", "Home", new { status = "success", eventId });
            }
            else
            {
                var errorDetails = response.Content;
                return View("Error", model: errorDetails);
            }
        }
    }
}
