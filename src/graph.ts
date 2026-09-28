import fetch from "node-fetch";
import { getAccessToken } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function graphGet<T>(path: string): Promise<T> {
  const token = await getAccessToken();
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph GET ${path} failed: ${res.status} ${body}`);
  }
  return (await res.json()) as T;
}

export interface GraphMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  webLink: string;
  receivedDateTime: string;
  from?: { emailAddress?: { name?: string; address?: string } };
}

export interface GraphEvent {
  id: string;
  subject: string;
  webLink: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  organizer?: { emailAddress?: { name?: string; address?: string } };
  isCancelled: boolean;
}

/** Messages received in the last N days, from the configured folder (default inbox). */
export async function fetchRecentMessages(lookbackDays: number): Promise<GraphMessage[]> {
  const folder = process.env.MS_MAIL_FOLDER ?? "inbox";
  const since = new Date(Date.now() - lookbackDays * 24 * 60 * 60 * 1000).toISOString();
  const select = "id,subject,bodyPreview,webLink,receivedDateTime,from";
  const filter = encodeURIComponent(`receivedDateTime ge ${since}`);
  const messages: GraphMessage[] = [];

  let path: string | undefined =
    `/me/mailFolders/${folder}/messages?$select=${select}&$filter=${filter}&$top=50&$orderby=receivedDateTime desc`;

  while (path) {
    const page: { value: GraphMessage[]; "@odata.nextLink"?: string } = await graphGet(path);
    messages.push(...page.value);
    path = page["@odata.nextLink"] ? page["@odata.nextLink"].replace(GRAPH_BASE, "") : undefined;
  }

  return messages;
}

/** Upcoming calendar events (i.e. accepted/received meeting invites) over the next N days. */
export async function fetchUpcomingEvents(lookForwardDays: number): Promise<GraphEvent[]> {
  const start = new Date().toISOString();
  const end = new Date(Date.now() + lookForwardDays * 24 * 60 * 60 * 1000).toISOString();
  const select = "id,subject,webLink,start,end,organizer,isCancelled";
  const path =
    `/me/calendarView?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}` +
    `&$select=${select}&$top=100&$orderby=start/dateTime`;

  const events: GraphEvent[] = [];
  let next: string | undefined = path;
  while (next) {
    const page: { value: GraphEvent[]; "@odata.nextLink"?: string } = await graphGet(next);
    events.push(...page.value.filter((e) => !e.isCancelled));
    next = page["@odata.nextLink"] ? page["@odata.nextLink"].replace(GRAPH_BASE, "") : undefined;
  }
  return events;
}
