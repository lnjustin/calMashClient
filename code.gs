function buildCard() {
  var card = CardService.newCardBuilder();
  
  var section = CardService.newCardSection();
  
  section.addWidget(CardService.newTextInput()
    .setFieldName("calendarUrl")
    .setTitle("Target Google Calendar Share Link"));
  
  section.addWidget(CardService.newTextInput()
    .setFieldName("calendarGroup")
    .setTitle("Calendar Group to Filter"));
  
  section.addWidget(CardService.newTextInput()
    .setFieldName("syncFrequency")
    .setTitle("Sync Frequency (in hours)"));
  
  section.addWidget(CardService.newTextButton()
    .setText("Save Settings")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("saveSettings")));
  
  section.addWidget(CardService.newTextButton()
    .setText("Sync Now")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("syncNow")));
  
  card.addSection(section);
  return card.build();
}

function saveSettings(e) {
  var userProps = PropertiesService.getUserProperties();
  var calendarUrl = e.formInput.calendarUrl;
  var calendarGroup = e.formInput.calendarGroup;
  var syncFrequency = e.formInput.syncFrequency;
  
  userProps.setProperty("calendarUrl", calendarUrl);
  userProps.setProperty("calendarGroup", calendarGroup);
  userProps.setProperty("syncFrequency", syncFrequency);
  
  var calendarId = extractCalendarId(calendarUrl);
  if (!calendarId) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Invalid Calendar URL."))
      .build();
  }
  
  var secondaryCalendar = createOrGetSecondaryCalendar(calendarId);
  scheduleSync(syncFrequency);
  syncNow();
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Settings saved and initial sync started."))
    .build();
}

function syncNow() {
  var userProps = PropertiesService.getUserProperties();
  var calendarUrl = userProps.getProperty("calendarUrl");
  var calendarGroup = userProps.getProperty("calendarGroup");
  var calendarId = extractCalendarId(calendarUrl);
  var secondaryCalendar = createOrGetSecondaryCalendar(calendarId);
  
  var events = Calendar.Events.list(calendarId, {timeMin: new Date().toISOString()}).items;
  var existingEventsMap = {};
  
  secondaryCalendar.getEvents(new Date(), new Date(new Date().setFullYear(new Date().getFullYear() + 1)))
    .forEach(function(event) {
      var eventHash = event.extendedProperties && event.extendedProperties.shared ? event.extendedProperties.shared.EventHash : null;
      if (eventHash) existingEventsMap[eventHash] = event;
    });
  
  events.forEach(function(event) {
    if (event.extendedProperties && event.extendedProperties.shared && event.extendedProperties.shared.CalendarGroup === calendarGroup) {
      return;
    }
    
    var eventHash = generateEventHash(event);
    if (!existingEventsMap[eventHash]) {
      var eventData = {
        summary: event.summary,
        description: event.description,
        location: event.location,
        start: event.start,
        end: event.end,
        colorId: event.colorId || "1",
        recurrence: event.recurrence || [],
        extendedProperties: { shared: { "EventHash": eventHash, "CalendarGroup": calendarGroup } }
      };
      Calendar.Events.insert(eventData, secondaryCalendar.getId());
    }
  });
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Sync completed successfully."))
    .build();
}

function extractCalendarId(calendarUrl) {
  var match = calendarUrl.match(/https:\/\/calendar\.google\.com\/calendar\/embed\?src=([^&]+)/);
  return match ? decodeURIComponent(match[1]) : null;
}

function createOrGetSecondaryCalendar(originalCalendarId) {
  var secondaryCalendarName = "Mirror of " + originalCalendarId;
  var calendars = Calendar.CalendarList.list().items;
  
  var existingCalendar = calendars.find(function(cal) { return cal.summary === secondaryCalendarName; });
  if (existingCalendar) {
    return CalendarApp.getCalendarById(existingCalendar.id);
  }
  
  return CalendarApp.createCalendar(secondaryCalendarName);
}

function scheduleSync(syncFrequency) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "syncNow") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger("syncNow")
    .timeBased()
    .everyHours(Math.max(1, parseInt(syncFrequency)))
    .create();
}

function generateEventHash(event) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, event.summary + event.start.dateTime + event.end.dateTime));
}
