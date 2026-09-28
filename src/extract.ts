import * as chrono from "chrono-node";
import type { GraphMessage } from "./graph.js";

const DEADLINE_KEYWORDS = [
  "deadline",
  "due by",
  "due date",
  "due on",
  "submit by",
  "respond by",
  "rsvp by",
  "closes on",
  "expires on",
  "last day to",
];

export interface ExtractedDate {
  type: "Deadline";
  title: string;
  date: Date;
  subject: string;
  sender: string;
  webLink: string;
  messageId: string;
  notes: string;
}

/** Finds likely deadline mentions in an email's subject + preview text. Only keeps future-facing dates. */
export function extractDeadlines(message: GraphMessage): ExtractedDate[] {
  const text = `${message.subject}\n${message.bodyPreview}`;
  const lowerText = text.toLowerCase();
  const hasDeadlineKeyword = DEADLINE_KEYWORDS.some((kw) => lowerText.includes(kw));
  if (!hasDeadlineKeyword) return [];

  const results = chrono.parse(text, message.receivedDateTime ? new Date(message.receivedDateTime) : undefined);
  const sender =
    message.from?.emailAddress?.name ?? message.from?.emailAddress?.address ?? "Unknown sender";

  return results
    .filter((r) => r.start.date() >= new Date(Date.now() - 24 * 60 * 60 * 1000))
    .map((r) => ({
      type: "Deadline" as const,
      title: message.subject || "(no subject)",
      date: r.start.date(),
      subject: message.subject,
      sender,
      webLink: message.webLink,
      messageId: message.id,
      notes: r.text,
    }));
}
