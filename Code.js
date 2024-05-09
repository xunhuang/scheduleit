const labelName = "ScheduleIt"; // Variable for the label name
const labelNameDone = "ScheduleIt_done"; // Variable for the label name
const defaultCalendarName = "ScheduleIt"; // Replace this with the name of your calendar
const scriptProperties = PropertiesService.getScriptProperties();
const openaiApiKey = scriptProperties.getProperty("OPENAI_KEY");
const TEST_MODE = false; // Set this to true for test mode, false for production
const SKIP_TAG_DONE_LABEL = true; // Set this to true for test mode, false for production

function getDateDaysAgo(days) {
  const dateDaysAgo = new Date();
  dateDaysAgo.setDate(dateDaysAgo.getDate() - days);
  return (
    (dateDaysAgo.getMonth() + 1).toString().padStart(2, "0") +
    "/" +
    dateDaysAgo.getDate().toString().padStart(2, "0") +
    "/" +
    dateDaysAgo.getFullYear()
  );
}

function createCalendarEventFromEmail() {
  const threeDaysAgo = getDateDaysAgo(30);

  Logger.log("Date three days ago: " + threeDaysAgo);
  const queries = [
    {
      id: "query1",
      query: `label:${labelName} -label:${labelNameDone}`,
      prefix: "[HRS] ",
      calendar: "Head Royce School",
    },
    // {
    //   id: "query2",
    //   query: `{list:alamedatroop7@googlegroups.com list:alamedatroop2@googlegroups.com} after:${threeDaysAgo} -label:${labelNameDone}`,
    //   prefix: "[BSA]",
    // },
    // {
    //   id: "query2",
    //   query: `{from:headroyce.org} after:${threeDaysAgo} -label:${labelNameDone}`,
    //   prefix: "[HRS]",
    //   calendar: "Head Royce School",
    // },
  ];
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
    const threads = GmailApp.search(queryObject.query);
    threads.forEach((thread) => {
      const messages = thread.getMessages();
      const message = messages[0];
      allMessages.push({
        message: message,
        queryId: queryObject.id,
        prefix: queryObject.prefix,
        calendar: queryObject.calendar, // Add the calendar field from the query object
      });
    });
  });

  allMessages.forEach((entry) => {
    // Determine the calendar to use: specific calendar from the query or the default one
    const calendarToUse = entry.calendar
      ? CalendarApp.getCalendarsByName(entry.calendar)[0]
      : defaultCalendar;
    if (!calendarToUse) {
      Logger.log("No calendar found with the name: " + entry.calendar);
      return;
    } else {
      Logger.log("Using Calendar with the name: " + calendarToUse.getName());
    }
    processMessage(entry.message, calendarToUse, entry.queryId, entry.prefix); // Pass the determined calendar to processMessage
  });
}

async function processMessage(message, calendar, queryId, prefix) {
  Logger.log(message.getSubject());
  const subject = message.getSubject();
  const content = message.getPlainBody();
  const emailUrl = `https://mail.google.com/mail/u/0/#inbox/${message.getId()}`;

  const events = await extractEventsUsingChatGPT(content, emailUrl);

  events.forEach((event) => {
    const { eventName, startTime, endTime } = event;
    if (!TEST_MODE) {
      calendar.createEvent(
        prefix + eventName,
        new Date(startTime),
        new Date(endTime),
        {
          description: `Extracted from email: ${emailUrl} \n ${content}`,
        }
      );
    }
    Logger.log(
      `Calendar event created: ${prefix}${eventName} from ${startTime} to ${endTime}`
    );
  });

  if (!SKIP_TAG_DONE_LABEL) {
    // Retrieve or create the "done" label
    const doneLabel =
      GmailApp.getUserLabelByName(labelNameDone) ||
      GmailApp.createLabel(labelNameDone);

    // Apply the "done" label to the thread containing the message
    const thread = message.getThread();
    thread.addLabel(doneLabel);
  }
}

// Function to call OpenAI's ChatGPT to extract event details
async function extractEventsUsingChatGPT(content, messageURL) {
  // Remove URLs from the content
  const contentWithoutUrls = content.replace(/(https?:\/\/[^\s]+)/g, "");
  Logger.log(contentWithoutUrls);
  const systemMessage = `
Please extract any calendar events including date, time, and event name from the following email 
and return the full set of data as a json array without duplicates, and do not truncate. 
If a year is not given, use the current year(2024), use ISO 8601 time format for start time assuming as default timezone. 
If no time is given, assume it is a full day event. 
Use startTime, endTime, eventName as fields name for the JSON.

If no events are found, return an empty JSON array.
`;
  try {
    const response = UrlFetchApp.fetch(
      "https://api.openai.com/v1/chat/completions",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${openaiApiKey}`,
          "Content-Type": "application/json",
        },
        payload: JSON.stringify({
          model: "gpt-4",
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
    const data = JSON.parse(text);
    const events = parseEventsFromResponse(data.choices[0].message.content);
    return events;
  } catch (e) {
    Logger.log(e);
    Logger.log("Errant message url: " + messageURL);
  }
  return [];
}

// Dummy function to parse events from ChatGPT response
function parseEventsFromResponse(response) {
  Logger.log(response);
  return JSON.parse(response);
}
