import Airtable from "airtable";

export interface InboxDateRecord {
  title: string;
  date: Date;
  type: "Event Invite" | "Deadline";
  sourceSubject: string;
  sender: string;
  emailLink: string;
  notes: string;
  outlookMessageId: string;
}

function getTable() {
  const apiKey = process.env.AIRTABLE_API_KEY;
  const baseId = process.env.AIRTABLE_BASE_ID;
  const tableName = process.env.AIRTABLE_TABLE_NAME ?? "Inbox Dates";
  if (!apiKey || !baseId) {
    throw new Error("AIRTABLE_API_KEY and AIRTABLE_BASE_ID must be set. Copy .env.example to .env and fill it in.");
  }
  const base = new Airtable({ apiKey }).base(baseId);
  return base(tableName);
}

/** Existing Outlook message IDs already synced, so we don't create duplicate rows on repeat runs. */
export async function fetchSyncedMessageIds(): Promise<Set<string>> {
  const table = getTable();
  const ids = new Set<string>();
  await table
    .select({ fields: ["Outlook Message ID"] })
    .eachPage((records, next) => {
      for (const record of records) {
        const id = record.get("Outlook Message ID");
        if (typeof id === "string" && id) ids.add(id);
      }
      next();
    });
  return ids;
}

export async function createRecords(records: InboxDateRecord[]): Promise<void> {
  const table = getTable();
  const chunkSize = 10; // Airtable's create limit per request
  for (let i = 0; i < records.length; i += chunkSize) {
    const chunk = records.slice(i, i + chunkSize).map((r) => ({
      fields: {
        Title: r.title,
        Date: r.date.toISOString().slice(0, 10),
        Type: r.type,
        "Source Subject": r.sourceSubject,
        Sender: r.sender,
        "Email Link": r.emailLink,
        Notes: r.notes,
        "Synced At": new Date().toISOString(),
        "Outlook Message ID": r.outlookMessageId,
      },
    }));
    await table.create(chunk);
  }
}
