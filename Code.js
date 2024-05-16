const labelName = "ScheduleIt";
const labelNameDone = "ScheduleIt_done";
const labelNameError = "ScheduleIt_error";
const labelNameNoEvent = "ScheduleIt_no_event";
const labelNameEventFound = "ScheduleIt_events_found";
const defaultCalendarName = "ScheduleIt";

const scriptProperties = PropertiesService.getScriptProperties();
const openaiApiKey = scriptProperties.getProperty("OPENAI_KEY");

const TEST_MODE = false;
const SKIP_TAG_DONE_LABEL = false;

function getDateDaysAgo(days) {
  const dateDaysAgo = new Date();
  dateDaysAgo.setDate(dateDaysAgo.getDate() - days);
  return dateDaysAgo.toLocaleDateString("en-US", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
}

function createCalendarEventFromEmail() {
  const lookbackPeriod = getDateDaysAgo(35);

  Logger.log("Reading Messages Since " + lookbackPeriod);
  var queries = [
    {
      query: `{from:jaszha2020@gmail.com from:scouting.org from:mailsrv@troopkit.com list:alamedatroop7@googlegroups.com list:alamedatroop2@googlegroups.com}`,
      prefix: "[BSA]",
      calendar: "Scouting",
    },
    {
      query: `{from:headroyce.org, from:ljackson@alumni.princeton.edu} `,
      prefix: "[HRS]",
      calendar: "Head Royce School",
    },
    {
      query: `"Encinal Jr-Sr High School" `,
      prefix: "[Encinal]",
      calendar: "Encinal",
    },
    {
      query: `"The Berkeley Chess School" `,
      prefix: "[chess]",
      calendar: "chess",
    },
    {
      query: `Subject:BMC `,
      prefix: "[BMC]",
      calendar: "BMC",
    },
    {
      query: `label:${labelName} -label:${labelNameDone}`,
    },
  ];
  if (TEST_MODE) {
    queries = [
      {
        query: `label:${labelName}`,
      },
    ];
  }
  const defaultCalendar =
    CalendarApp.getCalendarsByName(defaultCalendarName)[0];
  if (!defaultCalendar) {
    Logger.log(
      "No default calendar found with the name: " + defaultCalendarName
    );
    return;
  }

  const allMessages = [];

  queries.forEach((queryObject) => {
    var query = queryObject.query;
    if (!TEST_MODE) {
      query = query + " after:${lookbackPeriod} -label:${labelNameDone}";
    }
    Logger.log("Searching for: " + query);

    const threads = GmailApp.search(query);
    threads.forEach((thread) => {
      const messages = thread.getMessages();
      const message = messages[0];
      allMessages.push({
        message: message,
        prefix: queryObject.prefix,
        calendar: queryObject.calendar, // Add the calendar field from the query object
      });
    });
  });

  const processedMessageIds = new Set(); // Initialize set for processed message IDs
  allMessages.forEach((entry) => {
    if (processedMessageIds.has(entry.message.getId())) {
      Logger.log("Message already processed: " + entry.message.getId());
      return; // Skip processing if message already processed
    }

    processedMessageIds.add(entry.message.getId()); // Add message ID to set

    const calendarToUse = entry.calendar
      ? CalendarApp.getCalendarsByName(entry.calendar)[0]
      : defaultCalendar;

    if (!calendarToUse) {
      if (entry.calendar) {
        Logger.log(
          "No calendar found with the name: " +
            entry.calendar +
            ". Creating it."
        );
        const newCalendar = CalendarApp.createCalendar(entry.calendar);
        processMessage(entry.message, newCalendar, entry.prefix);
      } else {
        Logger.log(
          "No default calendar found with the name: " + defaultCalendarName
        );
        return;
      }
    } else {
      Logger.log("Using Calendar with the name: " + calendarToUse.getName());
      processMessage(entry.message, calendarToUse, entry.prefix);
    }
  });
}

async function processMessage(message, calendar, prefix) {
  Logger.log(message.getSubject());
  const content = message.getPlainBody();
  const timedate = message.getDate();
  const emailUrl = `https://mail.google.com/mail/u/0/#inbox/${message.getId()}`;

  try {
    const events = await extractEventsUsingChatGPT(
      `Message recieved on date: ${timedate} \n` + content,
      emailUrl
    );

    if (events.length === 0 || !events) {
      Logger.log("No events extracted. Exiting for: " + message.getSubject());
      applyLabelToMessage(message, labelNameNoEvent);
    } else {
      applyLabelToMessage(message, labelNameEventFound);
      events.forEach((event) => {
        const { eventName, startTime, endTime, fullDay } = event;
        const startDate = new Date(startTime);
        const endDate = new Date(endTime);
        const durationHours =
          (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60);
        const allDay = durationHours >= 20 || fullDay;

        const duplicateEvent = isDuplicateEvent(
          startDate,
          endDate,
          calendar,
          prefix,
          eventName
        );

        if (duplicateEvent) {
          Logger.log(`Duplicate event found: ${prefix}${eventName}`);
          return; // Skip this iteration
        }
        if (!TEST_MODE) {
          if (allDay) {
            // Create all-day event
            calendar.createAllDayEvent(prefix + eventName, startDate, {
              description: `Extracted from email: ${emailUrl} \n ${content}`,
            });
          } else {
            // Create event with specific start and end times
            calendar.createEvent(prefix + eventName, startDate, endDate, {
              description: `Extracted from email: ${emailUrl} \n ${content}`,
            });
          }
        }
        Logger.log(
          `Calendar event created: ${prefix}${eventName} from ${startTime} to ${endTime}`
        );
      });
    }
  } catch (e) {
    Logger.log("Error processing message: " + e);
    Logger.log("Errant message url: " + messageURL);
    applyLabelToMessage(message, labelNameError);
    removeLabelToMessage(message, labelNameDone);
    return;
  }

  if (!SKIP_TAG_DONE_LABEL) {
    // Retrieve or create the "done" label
    applyLabelToMessage(message, labelNameDone);
  }
}

function isDuplicateEvent(startDate, endDate, calendar, prefix, eventName) {
  const existingEvents = calendar.getEvents(startDate, endDate);
  Logger.log("Checking for dup: " + prefix + eventName);
  return existingEvents.find((existingEvent) => {
    Logger.log(existingEvent.getTitle());
    const similarity = jaroWinkler(
      existingEvent.getTitle(),
      prefix + eventName
    );
    return (
      similarity > 0.8 && // Adjust this threshold as needed
      existingEvent.getStartTime().getTime() === startDate.getTime()
    );
  });
}

function applyLabelToMessage(message, label) {
  const doneLabel =
    GmailApp.getUserLabelByName(label) || GmailApp.createLabel(label);
  const thread = message.getThread();
  thread.addLabel(doneLabel);
}

function removeLabelToMessage(message, label) {
  const doneLabel =
    GmailApp.getUserLabelByName(label) || GmailApp.createLabel(label);
  const thread = message.getThread();
  thread.removeLabel(doneLabel);
}

// Function to call OpenAI's ChatGPT to extract event details
async function extractEventsUsingChatGPT(content, messageURL) {
  // Remove URLs from the content
  const contentWithoutUrls = content.replace(/(https?:\/\/[^\s]+)/g, "");
  Logger.log(contentWithoutUrls);
  const currentYear = new Date().getFullYear();

  const systemMessage = `
Please extract any calendar events including date, time, and event name from the following email 
and return the full set of data as a json array without duplicates, and do not truncate. 
Use startTime, endTime, eventName, fullDay as required fields name for the JSON.

If a year is not given, deduce the event year based on the date email was recieved,
and if that's not obvious, use the current year(${currentYear}). 
Use ISO 8601 time format for start time assuming PST as default timezone. 
If no time is given, assume it is a full day event. 
fullDay should be a boolean value.

If no events are found, return an empty JSON array.

please return pure json string without any formatting or tags
`;
  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${openaiApiKey}`,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify({
        // model: "gpt-4",
        model: "gpt-4o",
        messages: [
          {
            role: "system",
            content: systemMessage,
          },
          { role: "user", content: contentWithoutUrls },
        ],
      }),
    }
  );

  const text = response.getContentText();
  Logger.log(text);
  const data = JSON.parse(text);
  const events = parseEventsFromResponse(data.choices[0].message.content);
  return events;
}

// Dummy function to parse events from ChatGPT response
function parseEventsFromResponse(response) {
  // Logger.log(response);
  try {
    return JSON.parse(response);
  } catch (e) {
    Logger.log("Error parsing response (JSON expectecd): " + e);
    return [];
  }
}

function jaroWinkler(s1, s2) {
  var m = 0;

  // Exit early if either are empty.
  if (s1.length === 0 || s2.length === 0) {
    return 0;
  }

  // Exit early if they're an exact match.
  if (s1 === s2) {
    return 1;
  }

  var range = Math.floor(Math.max(s1.length, s2.length) / 2) - 1,
    s1Matches = new Array(s1.length),
    s2Matches = new Array(s2.length);

  for (i = 0; i < s1.length; i++) {
    var low = i >= range ? i - range : 0,
      high = i + range <= s2.length - 1 ? i + range : s2.length - 1;

    for (j = low; j <= high; j++) {
      if (s1Matches[i] !== true && s2Matches[j] !== true && s1[i] === s2[j]) {
        ++m;
        s1Matches[i] = s2Matches[j] = true;
        break;
      }
    }
  }

  // If no matches were found, then we have a Jaro distance of 0.
  if (m === 0) {
    return 0;
  }

  // Count the transpositions.
  var k = (n_trans = 0);

  for (i = 0; i < s1.length; i++) {
    if (s1Matches[i] === true) {
      for (j = k; j < s2.length; j++) {
        if (s2Matches[j] === true) {
          k = j + 1;
          break;
        }
      }

      if (s1[i] !== s2[j]) {
        ++n_trans;
      }
    }
  }

  var weight = (m / s1.length + m / s2.length + (m - n_trans / 2) / m) / 3,
    l = 0,
    p = 0.1;

  if (weight > 0.7) {
    while (s1[l] === s2[l] && l < 4) {
      ++l;
    }

    weight = weight + l * p * (1 - weight);
  }

  return weight;
}
