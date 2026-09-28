import "dotenv/config";
import { fetchRecentMessages, fetchUpcomingEvents } from "./graph.js";
import { extractDeadlines } from "./extract.js";
import { createRecords, fetchSyncedMessageIds, type InboxDateRecord } from "./airtable.js";

async function main() {
  const lookbackDays = Number(process.env.LOOKBACK_DAYS ?? "14");

  console.log(`Fetching messages from the last ${lookbackDays} day(s)...`);
  const messages = await fetchRecentMessages(lookbackDays);
  console.log(`Fetching calendar events for the next ${lookbackDays} day(s)...`);
  const events = await fetchUpcomingEvents(lookbackDays);

  console.log("Checking which items are already synced...");
  const syncedIds = await fetchSyncedMessageIds();

  const records: InboxDateRecord[] = [];

  for (const event of events) {
    if (syncedIds.has(event.id)) continue;
    records.push({
      title: event.subject || "(no subject)",
      date: new Date(event.start.dateTime + "Z"),
      type: "Event Invite",
      sourceSubject: event.subject,
      sender:
        event.organizer?.emailAddress?.name ?? event.organizer?.emailAddress?.address ?? "Unknown organizer",
      emailLink: event.webLink,
      notes: "",
      outlookMessageId: event.id,
    });
  }

  for (const message of messages) {
    if (syncedIds.has(message.id)) continue;
    for (const deadline of extractDeadlines(message)) {
      records.push({
        title: deadline.title,
        date: deadline.date,
        type: "Deadline",
        sourceSubject: deadline.subject,
        sender: deadline.sender,
        emailLink: deadline.webLink,
        notes: deadline.notes,
        outlookMessageId: deadline.messageId,
      });
    }
  }

  if (records.length === 0) {
    console.log("Nothing new to sync.");
    return;
  }

  console.log(`Creating ${records.length} new record(s) in Airtable...`);
  await createRecords(records);
  console.log("Done.");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
