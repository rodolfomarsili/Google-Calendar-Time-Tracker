function processEventsData(startDate, endDate) {
  try {

    if (!(startDate && endDate)) {
      // Handle missing parameters
      throw new Error(consoleMessages.missingParameters);
    }

    const mappedEvents = getFilteredEvents(startDate, endDate);
  

    //--------------------Save to Spreadsheet------------
      const sheetMappedEvents = mappedEvents.map(event => {
        return [
          event.id,
          event.category,
          event.name,
          event.userEmail,
          event.organizerEmail,
          event.startTime,
          event.endTime,
          event.workingHours / (1000 * 60 * 60),
          event.restingHours / (1000 * 60 * 60)
        ]
      })

      let spreadsheetId = spreadsheetsData.testSheet.spreadsheetId;
      let sheetName = spreadsheetsData.testSheet.sheetName;

      let rowsToInsert = sheetMappedEvents.length;
      let sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

      let idColumnIndex = getColumnNumberByHeader(sheet, "id");
      let dateColumnIndex = getColumnNumberByHeader(sheet, "start_time");

      sheet.insertRowsBefore(2, rowsToInsert); // Insert rows before row 2
      sheet.getRange(2, 1, rowsToInsert, sheetMappedEvents[0].length).setValues(sheetMappedEvents);
      sheet.getDataRange().removeDuplicates([idColumnIndex]).sort({column: dateColumnIndex, ascending: false});


  } catch (error) {
    console.error("Error in function:", error);
  }
}

function getFilteredEvents(startDate, endDate) {
  const userTimeZone = Calendar.Settings.get("timezone").value;
  const eventsData = getEventsData(startDate, endDate, userTimeZone);
  let filteredEvents = [];
  
  const userEmail = Session.getUser().getEmail(); 
  const categories = getScriptProperties();

  eventsData.dates.forEach(dayDate => {
    
    const { workingLocationEvents, focusTimeEvents, defaultEvents, outOfOfficeEvents } = classifyEvents(eventsData.events, dayDate, userTimeZone, userEmail);

    if (workingLocationEvents.length === 0) {
      console.log(consoleMessages.noWorkingLocation);
      return;
    }

    if (workingLocationEvents.length > 1) {
      console.log(consoleMessages.multipleWorkingLocation);
      return;
    }

    const workingHours = {
      start: new Date(workingLocationEvents[0].start.dateTime),
      end: new Date(workingLocationEvents[0].end.dateTime)
    };

    const processedFocusTimeEvents = processFocusTimeEvents(focusTimeEvents, defaultEvents, workingHours, categories, userEmail);
    const processedDefaultEvents = processDefaultEvents(defaultEvents, workingLocationEvents[0], workingHours, categories, userEmail).reverse();

    const events = [...processedDefaultEvents, ...processedFocusTimeEvents];

    if (events.length === 0) {
      console.log(consoleMessages.noScheduledEvents);
      return;
    }

    console.log(dayDate, events.length);
    filteredEvents.push(...events);
  });

  return filteredEvents;
}

function classifyEvents(events, dayDate, userTimeZone, userEmail) {
  let workingLocationEvents = [];
  let focusTimeEvents = [];
  let defaultEvents = [];
  let outOfOfficeEvents = [];

  events.forEach(event => {
    if (isSameDate(event, userTimeZone, dayDate) && !isAllDayEvent(event)) {
      switch (event.eventType) {
        case "workingLocation":
          if (isWorkingHoursEvent(event)) workingLocationEvents.push(event);
          break;
        case "focusTime":
          focusTimeEvents.push(event);
          break;
        case "default":
          if (isMeetingAccepted(event, userEmail)) defaultEvents.push(event);
          break;
        case "outOfOffice":
          outOfOfficeEvents.push(event);
          break;
      }
    }
  });

  return { workingLocationEvents, focusTimeEvents, defaultEvents, outOfOfficeEvents };
}

function processFocusTimeEvents(focusTimeEvents, defaultEvents, workingHours, categories, userEmail) {
  return focusTimeEvents.map(event => {
    const eventStartTime = new Date(event.start.dateTime);
    const eventEndTime = new Date(event.end.dateTime);
    const overlaps = findOverlappingEvents(event, defaultEvents);
    const freeTime = calculateFreeTimeIntervals(event, overlaps);

    const focusTimeEventsWorkingHours = freeTime.freeTimeIntervals.map(freeTime => {
      const eventStartTime = new Date(freeTime.start);
      const eventEndTime = new Date(freeTime.end);
      return calculateWorkingHours(eventStartTime, eventEndTime, workingHours.start, workingHours.end);
    });

    const totals = focusTimeEventsWorkingHours.reduce((acc, curr) => {
      acc.workingHours += curr.workingHours;
      acc.restingHours += curr.restingHours;
      return acc;
    }, { workingHours: 0, restingHours: 0 });

    
    if(totals.workingHours + totals.restingHours === 0) {
      return;
    }

    const colorId = event.colorId;
    const eventCategory = categories[colorId]?.category || "***Category Not Found***";
    const eventName = event.summary || consoleMessages.untitledEvent;

    return {
      id: event.id,
      name: eventName,
      organizerEmail: event.organizer.email,
      userEmail: userEmail,
      category: eventCategory,
      startTime: eventStartTime,
      endTime: eventEndTime,
      workingHours: totals.workingHours,
      restingHours: totals.restingHours
    };
  }).filter(Boolean);
}


function processDefaultEvents(defaultEvents, workingLocationEvent, workingHours, categories, userEmail) {
  return defaultEvents.map(event => {
    const eventStartTime = new Date(event.start.dateTime);
    const eventEndTime = new Date(event.end.dateTime);
    const calculatedWorkingHours = calculateWorkingHours(eventStartTime, eventEndTime, workingHours.start, workingHours.end);
    const workingHours2 = findOverlappingEvents(event, [workingLocationEvent])[0];

    const colorId = event.colorId;
    const eventCategory = categories[colorId]?.category || "***Category Not Found***";
    const eventName = event.summary || consoleMessages.untitledEvent;

    return {
      id: event.id,
      name: eventName,
      organizerEmail: event.organizer.email,
      userEmail: userEmail,
      category: eventCategory,
      startTime: eventStartTime,
      endTime: eventEndTime,
      workingHours: calculatedWorkingHours.workingHours,
      restingHours: calculatedWorkingHours.restingHours
    };
  });
}

function getEventsData(startDate, endDate, timeZone) {

  // Ensure Date objects for startDate and endDate
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
    throw new Error("Invalid input: startDate or endDate must be valid Date objects.");
  }

  // Validate input parameters
  if (!timeZone) {
    throw new Error("Missing required parameter: timeZone.");
  }

  const millisecondsInADay = 24 * 60 * 60 * 1000;
  const daysDifference = Math.floor((endDate.getTime() - startDate.getTime()) / millisecondsInADay );  

  let dates = [];
  for (let i = daysDifference; i >= 0; i--) {
    let dayDate = Utilities.formatDate(
      new Date(startDate.getTime() + i * millisecondsInADay),
      timeZone,
      "MM/dd/yyyy"
    );
    dates.push(dayDate);
  }

  // Format start and end times with timezone and UTC offset
  const formattedStartDate = Utilities.formatDate(
    new Date(startDate),
    timeZone,
    "yyyy-MM-dd'T'00:00:00XXX"
  );
  const formattedendDate = Utilities.formatDate(
    new Date(endDate),
    timeZone,
    "yyyy-MM-dd'T'23:59:59XXX"
  );

  // Build query arguments
  const args = {
    showDeleted: false, // Exclude deleted events
    singleEvents: true, // Only include single events (not recurring events)
    timeMin: formattedStartDate, // Start time of the query
    timeMax: formattedendDate, // End time of the query
    orderBy: "startTime", // Order events by start time
    eventTypes: ["default", "focusTime", "outOfOffice", "workingLocation"]
  };

  // Retrieve events from primary calendar
  const events = Calendar.Events.list("primary", args);

  return {
    dates: dates,
    events: events.items
  }
}


/**
 * Finds and retrieves the overlaps between a primary event and an array of other events.
 * @param {Object} primaryEvent - The primary event object with start and end times.
 * @param {Array} otherEvents - An array of other event objects.
 * @returns {Array} - An array of objects representing the overlaps between the primary event and other events.
 */
function findOverlappingEvents(primaryEvent, otherEvents) {
  let overlappingEvents = [];

  for (let i = 0; i < otherEvents.length; i++) {
    let currentEvent = otherEvents[i];

    const primaryStart = new Date(primaryEvent.start.dateTime).getTime();
    const primaryEnd = new Date(primaryEvent.end.dateTime).getTime();
    const currentStart = new Date(currentEvent.start.dateTime).getTime();
    const currentEnd = new Date(currentEvent.end.dateTime).getTime();
    
    if (primaryEnd <= currentStart || primaryStart >= currentEnd) {
      continue; // Skip to the next iteration if there is no overlap
    }

    const overlapStart = Math.max(primaryStart, currentStart);
    const overlapEnd = Math.min(primaryEnd, currentEnd);

    const overlapDuration = overlapEnd - overlapStart;
    const remainingDuration = primaryEnd - primaryStart - overlapDuration;

    const adjustedoverlapDuration = Math.max(0, Math.min(overlapDuration, primaryEnd - primaryStart));
    const thereIsOverlapping = primaryEnd <= currentStart || primaryStart >= currentEnd;
    const remainingDuration2 = thereIsOverlapping ? currentEnd - currentStart : currentEnd - currentStart - adjustedoverlapDuration;

    let overlapDetail = {
      primaryEventId: primaryEvent.id,
      secodaryEventId: currentEvent.id,
      primaryEventTitle: primaryEvent.summary || "***Untitled Primary Event***",
      secodaryEventTitle: currentEvent.summary || "***Untitled Other Event***",
      start: overlapStart,
      end: overlapEnd,
      overlapDuration: overlapDuration,
      remainingPrimaryEventDuration: remainingDuration
    };
    overlappingEvents.push(overlapDetail);
  }

  return overlappingEvents;
}

/**
 * Calculates the available free time intervals within a primary event's duration by excluding time ranges of overlapping events.
 * 
 * @param {Object} primaryEvent - Object representing the primary event with start and end times.
 * @param {Array} overlappingEvents - Array of events that overlap with the primary event, including start and end times.
 * @returns {Object} - Object containing an array of free time intervals and the total free time in milliseconds.
 */
function calculateFreeTimeIntervals(primaryEvent, overlappingEvents) {

  // Store free time intervals
  let freeTimeIntervals = [];

  // Convert primary event start and end times to milliseconds
  let primaryEventTimeRange = {
    start: new Date(primaryEvent.start.dateTime).getTime(),
    end: new Date(primaryEvent.end.dateTime).getTime()
  };

  // Initialize to track the total free time available
  let totalAvailableFreeTime = primaryEventTimeRange.end - primaryEventTimeRange.start;

  // Directly return if there are no overlapping events
  if (overlappingEvents.length === 0) {
    freeTimeIntervals.push({
      eventId: primaryEvent.id,
      eventSummary: primaryEvent.summary || "***Untitled Event***",
      freeTimeInterval: totalAvailableFreeTime,
      start: primaryEventTimeRange.start,
      end: primaryEventTimeRange.end
    });
    return {
      totalFreeTime: totalAvailableFreeTime,
      freeTimeIntervals: freeTimeIntervals
    };
  };

  totalAvailableFreeTime = 0; // Reset total free time as it will be recalculated

  // Process the last overlap to calculate remaining free time
  let lastOverlapFreeTime = {
    eventId: primaryEvent.id,
    eventSummary: primaryEvent.summary || "***Untitled Event***",
    freeTimeInterval: (primaryEventTimeRange.end - overlappingEvents[overlappingEvents.length - 1].end),
    start: overlappingEvents[overlappingEvents.length - 1].end,
    end: primaryEventTimeRange.end
  };
  freeTimeIntervals.unshift(lastOverlapFreeTime); // Add to the beginning of the array
  totalAvailableFreeTime += lastOverlapFreeTime.freeTimeInterval;

  // Calculate free time intervals between overlaps
  for (let i = overlappingEvents.length - 1; i > 0; i--) {
    let freeTimeInterval = {
      freeTimeInterval: (overlappingEvents[i].start - overlappingEvents[i - 1].end),
      start: overlappingEvents[i - 1].end,
      end: overlappingEvents[i].start
    };
    freeTimeIntervals.unshift(freeTimeInterval); // Add to the beginning for correct order
    totalAvailableFreeTime += freeTimeInterval.freeTimeInterval;
  }

  // Calculate free time before the first overlap
  let firstOverlapFreeTime = {
    eventId: primaryEvent.id,
    eventSummary: primaryEvent.summary || "***Untitled Event***",
    freeTimeInterval: (overlappingEvents[0].start - primaryEventTimeRange.start),
    start: primaryEventTimeRange.start,
    end: overlappingEvents[0].start
  };
  freeTimeIntervals.unshift(firstOverlapFreeTime); // Ensure chronological order
  totalAvailableFreeTime += firstOverlapFreeTime.freeTimeInterval;

  // Filter out non-positive intervals and return structured data
  return {
    totalFreeTime: totalAvailableFreeTime,
    freeTimeIntervals: freeTimeIntervals.filter(interval => interval.freeTimeInterval > 0)
  };
}

function calculateWorkingHours(eventStartTime, eventEndTime, workingHoursStartTime, workingHoursEndTime) {
  try {
    validateDates(eventStartTime, eventEndTime, workingHoursStartTime, workingHoursEndTime);

    const eventStart = eventStartTime.getTime();
    const eventEnd = eventEndTime.getTime();
    const workingHoursStart = workingHoursStartTime.getTime();
    const workingHoursEnd = workingHoursEndTime.getTime();

    const workingHours = Math.min(eventEnd, workingHoursEnd) - Math.max(eventStart, workingHoursStart);
    const adjustedWorkingHours = Math.max(0, Math.min(workingHours, workingHoursEnd - workingHoursStart));

    const eventIsOutsideWorkingHours = eventEnd <= workingHoursStartTime || eventStart >= workingHoursEnd;

    const restingHours = eventIsOutsideWorkingHours ? eventEnd - eventStart : eventEnd - eventStart - adjustedWorkingHours;

    return {
      workingHours: adjustedWorkingHours,
      restingHours: restingHours
    };
  } catch (error) {
    console.error("Error in calculateWorkingHours:", error);
    throw new Error("An error occurred while calculating working hours.");
  }
}


function isMeetingAccepted(event, userEmail) {
  try {
    if (userEmail === event.organizer.email) {
      return true; // Organizer is always considered as accepted
    }
    for (let i = 0; i < event.attendees.length; i++) {
      const attendee = event.attendees[i];
      if (attendee.email === userEmail && attendee.responseStatus === "accepted") {
        return true; // User has explicitly accepted the meeting
      }
    }
    return false; // User is not the organizer and has not accepted the meeting
  } catch (error) {
    throw new Error("An error occurred while determining meeting acceptance status.");
  }
}

function validateDates(...dates) {
  if (!dates.every((date) => date instanceof Date)) {
    throw new Error("All parameters must be Date objects.");
  }
}

function isSameDate(event, timezone, date) {
  try {
    let eventDate = Utilities.formatDate(
      new Date(event.start.dateTime),
      timezone,
      "MM/dd/yyyy"
    );
    return date === eventDate;
  } catch (error) {
    throw new Error("An error occurred while checking event date.");
  }
}


function isAllDayEvent(event) {
  return ("date" in event.start && "date" in event.end);
}


function isWorkingHoursEvent(event) {
  try {
    if ("summary" in event) {
      return event.summary.includes("Working Hours");
    }
    return false;
  } catch (error) {
    console.error("Error in isWorkingHoursEvent:", error);
    throw new Error("An error occurred while checking working hours event.");
  }
}

function getColumnNumberByHeader(sheet, columnHeader) {
  // Get the values of the first row (header row) in the sheet
  const headerRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
  
  // Find the index of the specified column header in the header row
  const columnNumber = headerRow.indexOf(columnHeader);

  // If the column header is found, return its 1-indexed position; otherwise, return undefined
  if (columnNumber !== -1) {
    return columnNumber + 1;
  }
}

const spreadsheetsData = {
  testSheet: {
    spreadsheetId: "17uOSNJsjM8o1wD5CJM3WDCozbBbjJIoEGiicOSc6T00",
    sheetName: "Registro"
  }
}

const admins = ["rodolfo.marsili@xertica.com"];

const categories = ["", "event", "presales", "freetime", "internal", "training"];

const predefinedCategories = {
  '1': { name: 'Lavender', category: '', classic: '#A4BDFC', modern: '#7986CB' },
  '2': { name: 'Sage', category: '', classic: '#7AE7BF', modern: '#33B679' },
  '3': { name: 'Grape', category: 'event', classic: '#DBADFF', modern: '#8E24AA' },
  '4': { name: 'Flamingo', category: '', classic: '#FF887C', modern: '#E67C73' },
  '5': { name: 'Banana', category: 'presales', classic: '#FBD75B', modern: '#F6BF26' },
  '6': { name: 'Tangerine', category: '', classic: '#FFB878', modern: '#F4511E' },
  '7': { name: 'Peacock', category: 'freetime', classic: '#46D6DB', modern: '#039BE5' },
  '8': { name: 'Graphite', category: '', classic: '#E1E1E1', modern: '#616161' },
  '9': { name: 'Blueberry', category: '', classic: '#5484ED', modern: '#3F51B5' },
  '10': { name: 'Basil', category: 'internal', classic: '#51B749', modern: '#0B8043' },
  '11': { name: 'Tomato', category: 'training', classic: '#DC2127', modern: '#D50000' }
};

const consoleMessages = {
  missingParameters: "Either startDate and endDate must be provided.",
  noWorkingLocation: "There is no Working Location set for this date",
  multipleWorkingLocation: "Only 1 working location event is allowed per day",
  noScheduledEvents: "There are no scheduled events for this date",
  calculateWorkingHoursError: "An error occurred while calculating working hours.",
  meetingAcceptanceError: "An error occurred while determining meeting acceptance status.",
  checkEventDateError: "An error occurred while checking event date.",
  workingHoursEventError: "An error occurred while checking working hours event.",
  untitledEvent: "***Untitled Event***"
};

function showSideBar() {
  const template = HtmlService.createTemplateFromFile("sidebar");
  const html = template.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function showColorsTable() {
  const userEmail = Session.getUser().getEmail();
  const scriptProperties = getScriptProperties();
  let template = HtmlService.createTemplateFromFile("colors_table");
  template.scriptProperties = JSON.stringify(scriptProperties);
  template.categories = JSON.stringify(categories);
  
  template.admins = JSON.stringify(admins);
  template.userEmail = JSON.stringify(userEmail);
  const html = template.evaluate()
                       .setWidth(550)
                       .setHeight(700);
                       
  SpreadsheetApp.getUi().showModalDialog(html, 'Colors Table');
}

function updateSheet(inputFrom, inputTo){
  const dateFrom = formatDateFromHTML(inputFrom);
  const dateTo = formatDateFromHTML(inputTo);
  cleanSheet();
  processEventsData(dateFrom, dateTo);
  return;
}

function updateScriptProperties(updatedColors) {
  const scriptProperties = PropertiesService.getScriptProperties();

  if(!updatedColors && Object.keys(scriptProperties.getKeys()).length === 0){
    scriptProperties.setProperties(predefinedCategories);
    return
  }
  
  scriptProperties.setProperties(updatedColors);
}

function getScriptProperties(){
  const scriptProperties = PropertiesService.getScriptProperties();

  if(Object.keys(scriptProperties.getKeys()).length === 0){
    return predefinedCategories;
  }
  return convertToObject(scriptProperties.getProperties());
}

function cleanSheet(){
  const sheet = SpreadsheetApp.openById(spreadsheetsData.testSheet.spreadsheetId).getSheetByName(spreadsheetsData.testSheet.sheetName);
  sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clear();
}


function convertToObject(input) {
  let obj = {};
  for (let key in input) {
    let value = input[key].replace(/[{}]/g, '');
    let pairs = value.split(', ');
    let tempObj = {};
    pairs.forEach(pair => {
      let [prop, val] = pair.split('=');
      tempObj[prop] = val;
    });

    obj[key] = tempObj;
  }
  return obj;
}

function formatDateFromHTML(date) {
  // Regular expression to validate date format
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  if (!dateRegex.test(date)) {
    throw new Error(`Invalid date format: "${date}". Expected YYYY-MM-DD.`);
  }
  const splitedDate = date.split("-");
  // Reverse order to create Date object in MM-DD-YYYY format
  const formattedDate = new Date(`${splitedDate[1]}/${splitedDate[2]}/${splitedDate[0]}`);
  return formattedDate;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
