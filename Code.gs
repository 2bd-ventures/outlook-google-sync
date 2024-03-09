function updateCalendars() {
  // Update
  var CALENDAR_ID = 'Place_Your_Calendar_ID';
  var ICAL_URL = 'Place_Your_Published_Outlook_Calendar_URL';
  syncCalendar(CALENDAR_ID, ICAL_URL);
}

function syncCalendar(googleCalendarId, outlookIcsUrl) {
  var outlookId = 'outlook-id'

  var calendar = CalendarApp.getCalendarById(googleCalendarId);
  var iCalEvents = parseICal(outlookIcsUrl);

  var today = new Date();
  var thirtyDaysLater = new Date();
  thirtyDaysLater.setDate(today.getDate() + 30)
  var googleEvents = calendar.getEvents(today, thirtyDaysLater); // Adjust date range as needed
  var googleEventMap = {};

  // Create a map of Google Calendar events for easy lookup
  googleEvents.forEach(function (event) {
    var id = event.getTag(outlookId);
    googleEventMap[id] = event;
  });

  // Process iCal events
  iCalEvents.forEach(function (iCalEvent) {
    var eventKey = iCalEvent.uid + '-' + iCalEvent.sequence;
    var googleEvent = googleEventMap[eventKey];

    if (!googleEvent) {
      // Event does not exist in Google Calendar, create it
      var ev = calendar.createEvent(iCalEvent.summary, iCalEvent.start, iCalEvent.end, {
        description: iCalEvent.description,
        guests: 'dummy@example.com'
      });

      ev.setTag(outlookId, eventKey);
    } else if ( (+new Date(iCalEvent.start) != +new Date(googleEvent.getStartTime())) || (+new Date(iCalEvent.end) != +new Date(googleEvent.getEndTime()))){
      // confirm the time remains the same, and adjust
      googleEvent.setTime(iCalEvent.start, iCalEvent.end);
    }

    // Remove the event from the map since it has been processed
    delete googleEventMap[eventKey];
  });


  // Any events left in the map are not present in the iCal feed, delete them
  for (var id in googleEventMap) {
    googleEventMap[id].deleteEvent();
  }
}

function parseICal(url) {
  var response = UrlFetchApp.fetch(url);
  var iCalData = response.getContentText();
  var events = [];
  var eventRegex = /BEGIN:VEVENT[\s\S]*?END:VEVENT/g;
  var matches = iCalData.match(eventRegex);
  var now = new Date(); // Get the current time

  if (matches) {
    matches.forEach(function (eventData) {
      // Preprocess eventData to handle multi-line fields
      var preprocessedEventData = eventData.replace(/\n /g, '')

      var event = {};
      var lines = preprocessedEventData.split('\n');

      lines.forEach(function (line) {
        if (line.startsWith('SUMMARY:')) {
          event.summary = line.substring(8);
        } else if (line.startsWith('DESCRIPTION:')) {
          event.description = line.substring(12);
        } else if (line.startsWith('DTSTART;')) {
          event.start = convertStringToDate(line);
        } else if (line.startsWith('DTEND;')) {
          event.end = convertStringToDate(line);
        } else if (line.startsWith('UID:')) {
          event.uid = line.substring(4).trim();
        } else if (line.startsWith('SEQUENCE:')) {
          event.sequence = parseInt(line.substring(9), 10);
        }
      });

      // Only include events that end in the future
      if (event.end > now) {
        events.push(event);
      }
    });
  }

  return events;
}

function convertStringToDate(inputString) {
  // Extract the date part of the string
  var datePart = inputString.match(/(\d{8})T(\d{6})/);

  // If the date part is not found, return null
  if (!datePart) {
    return null;
  }

  // Extract the year, month, day, hours, minutes, and seconds
  var year = parseInt(datePart[1].substr(0, 4), 10);
  var month = parseInt(datePart[1].substr(4, 2), 10) - 1; // JS months are 0-indexed
  var day = parseInt(datePart[1].substr(6, 2), 10);
  var hours = parseInt(datePart[2].substr(0, 2), 10);
  var minutes = parseInt(datePart[2].substr(2, 2), 10);
  var seconds = parseInt(datePart[2].substr(4, 2), 10);

  // Create a Date object in the GMT Standard Time zone
  var date = new Date(Date.UTC(year, month, day, hours, minutes, seconds));

  // Return the Date object
  return date;
}




