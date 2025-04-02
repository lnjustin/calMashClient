var defaultMaxRetries = 10; // Maximum number of retries for api functions (with exponential backoff)
// Syncing logic can set this to true to cause the Google Apps Script "Executions" dashboard to report failure

  // Icons for the buttons - use actual icon URLs
var doneIconUrl = "https://github.com/lnjustin/App-Images/blob/master/CalendarMashup/check.png?raw=true";
var syncIconUrl = "https://github.com/lnjustin/App-Images/blob/master/CalendarMashup/mashup.png?raw=true";
var clearIconUrl = "https://github.com/lnjustin/App-Images/blob/master/CalendarMashup/delete.png?raw=true";

function buildCard() {

  var scriptProperties = PropertiesService.getScriptProperties();
  var storedSourceCalendarUrl = scriptProperties.getProperty("sourceCalendarUrl") || "";
  var storedTargetCalendarName = scriptProperties.getProperty("targetCalendarName") || "";
  var storedFilteredCalendarGroups = scriptProperties.getProperty("calendarGroups") || "";
  var storedSyncFrequency = scriptProperties.getProperty("syncFrequency") || "";
  var currentSchedule = scriptProperties.getProperty("syncScheduled") || "No sync scheduled.";

  const cardHeader = CardService.newCardHeader()
    .setTitle('Calendar Mashup')
    .setImageStyle(CardService.ImageStyle.CIRCLE)
    .setImageUrl("https://github.com/lnjustin/App-Images/blob/master/CalendarMashup/mashup.png?raw=true");

  var card = CardService.newCardBuilder();
  
  var section = CardService.newCardSection();
  
  section.addWidget(CardService.newTextParagraph().setText(
      "The CalMash app creates a Google Calendar from multiple source calendars that are grouped into one or more calendar groups. This client companion app mirrors that Google Calendar to a new calendar, but filters out events sourced from one or more specified calendar groups."
    ))
  section.addWidget(CardService.newTextInput()
    .setFieldName("sourceCalendarUrl")
    .setValue(storedSourceCalendarUrl)
    .setTitle("Google Calendar Share Link")
    .setHint("The Google Calendar to mirror."))
  
  section.addWidget(CardService.newTextInput()
    .setFieldName("targetCalendarName")
    .setValue(storedTargetCalendarName)
    .setTitle("Enter Mirror Calendar Name")
    .setHint("The name of the calendar to which to mirror events. Calendar will be created if not already in existence."))

  section.addWidget(CardService.newTextInput()
    .setFieldName("filteredCalendarGroups")
    .setValue(storedFilteredCalendarGroups)
    .setTitle("Calendar Group(s) to Filter Out (comma-separated)"));
  
  section.addWidget(CardService.newTextInput()
    .setFieldName("syncFrequency")
    .setValue(storedSyncFrequency)
    .setTitle("Sync Frequency (in hours)"));
  
  section.addWidget(CardService.newTextParagraph().setText("<b>Current Sync Status:</b> " + currentSchedule))
  
  section.addWidget(CardService.newTextButton()
    .setText("Save Settings")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("saveSettings")));
  
  section.addWidget(CardService.newTextButton()
    .setText("Sync Now")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("syncNow")));
  
  card.setHeader(cardHeader);
  card.addSection(section);

  flushLogs();
  return card.build();
}

function saveSettings(e) {
  clearTimeBasedTriggers();

  var formInput = e.formInput || {};
  var sourceCalendarUrl = formInput["sourceCalendarUrl"] ? formInput["sourceCalendarUrl"].trim() : "";
  var targetCalendarName = formInput["targetCalendarName"] ? formInput["targetCalendarName"].trim() : "";
  var filteredCalendarGroups = formInput["filteredCalendarGroups"] ? formInput["filteredCalendarGroups"].trim() : "";
  var syncFrequency = formInput["syncFrequency"] ? formInput["syncFrequency"].trim() : "";
  
  var scriptProperties = PropertiesService.getScriptProperties();

  // Store updated properties
  scriptProperties.setProperty("sourceCalendarUrl", sourceCalendarUrl);
  scriptProperties.setProperty("targetCalendarName", targetCalendarName);
  scriptProperties.setProperty("filteredCalendarGroups", filteredCalendarGroups);
  scriptProperties.setProperty("syncFrequency", syncFrequency);
  
  var frequencyLabel = (syncFrequency === "1") ? "hour" : "hours";
  scriptProperties.setProperty("syncScheduled", "Sync scheduled every " + syncFrequency + " " + frequencyLabel);

  writeLog("Scheduling background save to start momentarily.");
  ScriptApp.newTrigger("backgroundSave")
    .timeBased()
    .after(1000) // Start in 1 second
    .create();
  flushLogs();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Settings Saved. Starting background setup and sync."))
    .build();
}

function getInputValue(formInputs, fieldName) {
  return formInputs[fieldName] && formInputs[fieldName].stringInputs ? formInputs[fieldName].stringInputs.value[0] || "" : "";
}


function backgroundSave() {
  writeLog("Starting background save now.");
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var sourceCalendarUrl = scriptProperties.getProperty("sourceCalendarUrl") || "";
    if (!sourceCalendarUrl) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("⚠️ Please enter a calendar url."))
        .build();
    }

    var sourceCalendarId = extractCalendarId(sourceCalendarUrl);
    if (!sourceCalendarId) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("⚠️ Invalid Calendar URL."))
        .build();
    }

    var sourceCalendar = Calendar.Calendars.get(sourceCalendarId);
    var sourceCalendarName = sourceCalendar && sourceCalendar.summary ? sourceCalendar.summary : ""

    // Create target calendar if needed
    var targetCalendar = createOrGetTargetCalendar(sourceCalendar);
    if (!targetCalendar) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("❌ Failed to create or retrieve mirror calendar."))
        .build();
    }

    // Schedule sync
    var scheduleSuccess = scheduleSync();
    if (!scheduleSuccess) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("⚠️ Calendar created, but scheduling of filtered sync failed."))
        .setNavigation(CardService.newNavigation().updateCard(onHomepage())) // Reload the card by calling onHomepage()
        .build();
    }

    try {
      sync(); // sync after settings created or updated
    } catch (error) {
      writeLog("Error trying to sync calendar: " + error.toString());
      return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("⚠️ Filtered Calendar Sync Failed."))
      .setNavigation(CardService.newNavigation().updateCard(onHomepage())) // Reload the card by calling onHomepage()
      .build();
    }
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("✅ CalMash Client setup success!"))
      .setNavigation(CardService.newNavigation().updateCard(onHomepage())) // Reload the card by calling onHomepage()
      .build();
  } catch (error) {
    writeLog("Error in background save: " + error.toString());

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("❌ Background save error. Check logs."))
      .build();
  } finally {
    flushLogs(); // Ensure logs are flushed no matter what
  }
}

function clearTimeBasedTriggers() {
      // Remove existing triggers to prevent duplicates
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "scheduledSync" || triggers[i].getHandlerFunction() === "backgroundSave" || triggers[i].getHandlerFunction() === "onDemandSync") {
        ScriptApp.deleteTrigger(triggers[i]);
        writeLog("Deleted existing time-based trigger.");
      }
    }
}

function scheduledSync() {
  sync();
  flushLogs();
}

function onDemandSync() {
  writeLog("Executing on-demand sync now.");
  sync();
  scheduleSync();
  flushLogs();
}

function syncNow() {
  try {    
    clearTimeBasedTriggers();
    ScriptApp.newTrigger("onDemandSync")
      .timeBased()
      .after(1000) // Start in 1 second
      .create();

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Background sync started."))
      .build();
  } catch (error) {
      writeLog("Error: " + error.message);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(error.message))
        .setNavigation(CardService.newNavigation().updateCard(onHomepage())) // Reload homepage
        .build();
  } finally {
    flushLogs(); // Ensure logs are flushed no matter what
  }
}

function sync() {
  try {
    var userProperties = PropertiesService.getUserProperties();

    if (userProperties.getProperty('LastRun') > 0 && (new Date().getTime() - userProperties.getProperty('LastRun')) < 360000) {
      writeLog("Another iteration is currently running! Exiting...");
      flushLogs();
      return;
    }
    userProperties.setProperty('LastRun', new Date().getTime());

    var sourceCalendarUrl = userProperties.getProperty("calendarUrl");
    var sourceCalendarId = extractCalendarId(sourceCalendarUrl);
    var filteredCalendarGroupsString = userProperties.getProperty("filteredCalendarGroups");
    var filteredCalendarGroups = filteredCalendarGroupsString.split(",");
    // Trim whitespace from each group name using a traditional loop
    for (var i = 0; i < filteredCalendarGroups.length; i++) {
      filteredCalendarGroups[i] = filteredCalendarGroups[i].trim();
    }
    
    var targetCalendar = createOrGetTargetCalendar(sourceCalendarId);
    var existingTargetEvents = getTargetCalendarEvents(targetCalendar, null, null)
    var existingTargetEventsMap = {};
    existingTargetEvents.forEach(function(event) {
      var eventHash = event.extendedProperties && event.extendedProperties.shared && event.extendedProperties.shared.EventHash ? event.extendedProperties.shared.EventHash : null;
      if (eventHash) existingTargetEventsMap[eventHash] = event;
    });

    var syncedEventHashes = {};
    var sourceEvents = Calendar.Events.list(sourceCalendarId, {timeMin: new Date().toISOString()}).items;
    sourceEvents.forEach(function(sourceEvent) {
      var doMirrorEvent = false;
      var sourceEventHash;
      var sourceEventCalendarGroup;
      if (!sourceEvent.extendedProperties || !sourceEvent.extendedProperties.shared || !sourceEvent.extendedProperties.shared.CalendarGroup) { 
        // sync if source event does not have any calendar group (e.g., manually added to source calendar)
        doMirrorEvent = true; 
        sourceEventHash = generateEventHash(sourceEvent);
      }
      else {
        sourceEventCalendarGroup = sourceEvent.extendedProperties.shared.CalendarGroup;
        if (filteredCalendarGroups.indexOf(sourceEventCalendarGroup) === -1) {
          // sync if calendar group is not included in the list of groups to be filtered out
          doMirrorEvent = true;
          if (sourceEvent.extendedProperties.shared.EventHash) {
            sourceEventHash = sourceEvent.extendedProperties.shared.EventHash;
          }
          else sourceEventHash = generateEventHash(sourceEvent);
        }
      }
      if (doMirrorEvent) {
        syncedEventHashes[sourceEventHash] = true;
        if (!existingTargetEventsMap[sourceEventHash]) {
          // if target calendar does not already have the source event, add it
          var targetEventData = {
            summary: sourceEvent.summary,
            description: sourceEvent.description,
            location: sourceEvent.location,
            start: sourceEvent.start,
            end: sourceEvent.end,
            colorId: sourceEvent.colorId || "1",
            recurrence: sourceEvent.recurrence || [],
            extendedProperties: { shared: { "EventHash": sourceEventHash } }
          };
          if (sourceEventCalendarGroup) {
            targetEventData.extendedProperties.shared.CalendarGroup = sourceEventCalendarGroup;
          }
          callWithBackoff(function() {
            Calendar.Events.insert(targetEventData, targetCalendar.getId());
            writeLog("Added new event: " + sourceEvent.summary);
          }, defaultMaxRetries);
        }
      }
    });
    
    // Remove events no longer in source calendar
    existingTargetEvents.forEach(function(event) {
      var extendedProps = event.extendedProperties ? event.extendedProperties.shared : {};
      var eventHash = extendedProps["EventHash"];
      if (eventHash && !syncedEventHashes[eventHash]) {
        callWithBackoff(function() {
          Calendar.Events.remove(targetCalendar.getId(), event.id);
          writeLog("Deleted old/modified event: " + event.summary);
        }, defaultMaxRetries);
      }
    });

    writeLog("Sync finished!");
    return notifyUser("Mashup complete!");
  }
  catch (error) {
      writeLog("Error: " + error.message);

      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(error.message))
        .setNavigation(CardService.newNavigation().updateCard(onHomepage())) // Reload homepage
        .build();
  } finally {
    userProperties.setProperty('LastRun', 0);
    flushLogs(); // Ensure logs are flushed no matter what
  }
}

function notifyUser(message) {
  return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText(message)).build();
}

function getTargetCalendarEvents(targetCalendar, minTime, maxTime) {
  var existingEvents = [];
  var pageToken;
  do {
    var response = callWithBackoff(function() {
      return Calendar.Events.list(targetCalendar.getId(), {
        timeMin: minTime,
        timeMax: maxTime,
        maxResults: 250,
        pageToken: pageToken
      });
    }, defaultMaxRetries);
    if (response.items) {
      existingEvents = existingEvents.concat(response.items);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return existingEvents;
}

function extractCalendarId(calendarUrl) {
  var match = calendarUrl.match(/https:\/\/calendar\.google\.com\/calendar\/embed\?src=([^&]+)/);
  return match ? decodeURIComponent(match[1]) : null;
}

function createOrGetTargetCalendar(sourceCalendarName) {
  var targetCalendarName = sourceCalendarName + " " + "Mirror";
  var scriptProperties = PropertiesService.getScriptProperties();
  var oldCalendarId = scriptProperties.getProperty("targetCalendarId");

  var targetCalendar;
  
  if (oldCalendarId) {
    try {
      // Check if the calendar exists using the Advanced Calendar API
      targetCalendar = Calendar.Calendars.get(oldCalendarId);
      if (targetCalendar && targetCalendar.summary !== targetCalendarName) {
        Calendar.Calendars.patch({ summary: targetCalendarName }, oldCalendarId);
        writeLog("Renamed calendar to " + targetCalendarName);
      }
    } catch (e) {
      writeLog("Caught error: " + e.toString());
      writeLog("Stored calendar ID not found. Creating a new one.");
      targetCalendar = Calendar.Calendars.insert({ summary: targetCalendarName, timeZone: Session.getScriptTimeZone() });
    }
  } else {
    targetCalendar = Calendar.Calendars.insert({ summary: targetCalendarName, timeZone: Session.getScriptTimeZone() });
  }

  scriptProperties.setProperty("targetCalendarId", targetCalendar.id);
  scriptProperties.setProperty("targetCalendarName", targetCalendarName);

  return targetCalendar;
}

function scheduleSync() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var syncFrequency = scriptProperties.getProperty("syncFrequency") || "";
  var hoursInterval = Math.max(1, parseInt(syncFrequency)); // Ensure minimum 1-hour interval
  var triggers = ScriptApp.getProjectTriggers();
  clearTimeBasedTriggers();
  
  ScriptApp.newTrigger("scheduledSync")
    .timeBased()
    .everyHours(hoursInterval)
    .create();
}

function generateEventHash(event) {
    function normalizeDateTime(dateTimeObj) {
       // writeLog("Normalizing Datetime: " + dateTimeObj.dateTime)
        if (!dateTimeObj || !dateTimeObj.dateTime) return '';
        var date = new Date(dateTimeObj.dateTime); // Convert to Date object
        return date.toISOString(); // Normalize to UTC string
    }

    var eventTitle = event.title || '';
    var eventLocation = event.location ? event.location.trim() : ''; // Ensure location is considered
    var startTime = normalizeDateTime(event.startTime);
    var endTime = normalizeDateTime(event.endTime);

    var hashInput = eventTitle + '-' + eventLocation + '-' + startTime + '-' + endTime;

    var eventHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, hashInput));
    return eventHash;
}


function callWithBackoff(func, maxRetries) {
  var tries = 0;
  var result;
  while ( tries <= maxRetries ) {
    tries++;
    try{
      result = func();
      return result;
    }
    catch(err){
      err = err.message  || err;
      if ( err.indexOf("is not a function") !== -1  || !recoverableErrors.some(function(e){
              return err.toLowerCase().indexOf(e) !== -1;
            }) ) {
        throw err;
      } else if ( tries > maxRetries) {
        writeLog("Error, giving up after trying ${maxRetries} times [${err}]");
        return null;
      } else {
        writeLog( "Error, Retrying... [" + err  +"]");
        Utilities.sleep (Math.pow(2,tries)*100) + (Math.round(Math.random() * 100));
      }
    }
  }
  return null;
}

var logBuffer = [];

function writeLog(message) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var storedDebugLogging = scriptProperties.getProperty("debugLogging") === "true";
  if (storedDebugLogging) {
    logBuffer.push(message);
    if (logBuffer.length >= 20) {  // Write logs in chunks of 20
      flushLogs();
    }
  }
}

function flushLogs() {
  if (logBuffer.length > 0) {
    console.log(logBuffer.join("\n"));
    logBuffer = []; // Clear the buffer after flushing
  }
}


var recoverableErrors = [
  "service invoked too many times in a short time",
  "too many calendars or calendar events in a short time",
  "rate limit exceeded",
  "internal error",
  "http error 403", // forbidden
  "http error 408", // request timeout
  "http error 423", // locked
  "http error 500", // internal server error
  "http error 503", // service unavailable
  "http error 504"  // gateway timeout
];
