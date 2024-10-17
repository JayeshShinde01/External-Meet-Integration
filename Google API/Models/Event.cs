namespace GoogleAPI.Models
{
    public class Event
    {
        public string Summary { get; set; } // The title of the event
        public string Description { get; set; } // The description of the event
        public EventDateTime Start { get; set; } // Start date-time and time zone
        public EventDateTime End { get; set; } // End date-time and time zone
        public string Location { get; set; } // External meeting link (e.g., Zoom, Teams, etc.)
        public ConferenceData ConferenceData { get; set; } // Conference data for Google Meet (if applicable)
    }

    // Class for DateTime and TimeZone
    public class EventDateTime
    {
        public string DateTime { get; set; } // The date and time in "yyyy-MM-dd'T'HH:mm:ss.fffK" format
        public string TimeZone { get; set; } // The time zone (e.g., "America/New_York")
    }

    // Class for Conference Data (Google Meet)
    public class ConferenceData
    {
        public CreateConferenceRequest CreateRequest { get; set; }
        public string ConferenceId { get; internal set; }
    }

    // Class for creating the conference request (Google Meet)
    public class CreateConferenceRequest
    {
        public string RequestId { get; set; } // A unique request ID to avoid duplicate requests
        public ConferenceSolutionKey ConferenceSolutionKey { get; set; } // Specifies the solution key for Google Meet
    }

    // Specifies the type of conference solution (Google Meet)
    public class ConferenceSolutionKey
    {
        public string Type { get; set; } // Should be "hangoutsMeet" to create Google Meet link
    }
}
