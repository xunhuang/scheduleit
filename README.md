# scheduleit

This project is a Google Apps Script that will allow you to create calendar events from emails.

- It creates calendar if one doesn't exist
- It creates calendar events from email content. It uses OpenAI's API to extract the events
- It attempts deduplication for events
- It attempts to distinguish full-day event from normal events

## Installation

1. Create an Apps Script Project at https://script.google.com/home
2. Copy Code.js to the project
3. Replace openaiApiKey variable with key for OpenAI, in the form of sk-xxxxx
4. Configure UserRules to match your needs
   4.1 You can test your query by typing it into gmail's search bar and seeing if it returns the emails you expect. Those emails will be scanned for events
5. You can run the project manually, via this function "createCalendarEventFromEmail"
6. When you are happy with results, configure it to run on a schedule, by creating a (Time-based) trigger. My seeting is every 15 minutes.

## Future Experiment

- Use different LLMs (like llama 3 or Claude)
- Use LLM for better dedup compared to string similarity
