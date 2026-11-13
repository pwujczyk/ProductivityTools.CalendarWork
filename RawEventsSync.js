// Minimal raw calendar events sync -> ProductivityTools.CalendarLog.Api
// Existing scripts (Runner.js, Analysis.js, Conversion.js, Menu.js) are untouched.
// Uses Google Calendar API v3 Advanced Service (enabled in appsscript.json).
// No parsing, mapping, rule logic, or filtering happens here:
// - Raw Google Calendar v3 items are POSTed directly to the API.
// - Server-computed mutation commands (MOVE / RENAME) are executed blindly via Calendar API v3.

const RAW_SYNC_API_URL = "https://calendarlog-api.productivitytools.top";

/**
 * Main entry point (for time-driven triggers or manual runs):
 * 1. GET /api/sync/plan?account=... (server decides calendars & startDate/endDate)
 * 2. Calendar.Events.list (singleEvents=true)
 * 3. POST /api/sync/raw-events with raw v3 event objects
 * 4. Execute any server-returned mutations (MOVE / RENAME) and resync touched calendars once
 */
function syncRawCalendarEvents() {
  var account = rawSyncGetAccount_();
  var plan = rawSyncFetchPlan_(account);
  var calendars = plan.calendars || [];
  if (calendars.length === 0) {
    console.warn("API returned no active calendars for account:", account);
    return {
      account: account,
      calendarsSynced: 0,
      calendarsFailed: 0,
      totalEventsSent: 0,
      mutationsApplied: 0,
      mutationsFailed: 0,
      calendarsResynced: 0
    };
  }

  return rawSyncExecuteCalendarPlans_(account, calendars);
}

/**
 * Manual helper to sync a specific date range (inclusive) for all active calendars returned by the API.
 * Example: syncRawCalendarEventsForRange("2026-10-01", "2026-11-16")
 */
function syncRawCalendarEventsForRange(startDateStr, endDateStr) {
  var account = rawSyncGetAccount_();
  var plan = rawSyncFetchPlan_(account);
  var tz = Session.getScriptTimeZone();
  var effectiveEnd = endDateStr || Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var calendars = (plan.calendars || []).map(function (c) {
    return { calendarId: c.calendarId, startDate: startDateStr, endDate: effectiveEnd };
  });
  if (calendars.length === 0) {
    console.warn("API returned no active calendars for account:", account);
    return {
      account: account,
      calendarsSynced: 0,
      calendarsFailed: 0,
      totalEventsSent: 0,
      mutationsApplied: 0,
      mutationsFailed: 0,
      calendarsResynced: 0
    };
  }

  return rawSyncExecuteCalendarPlans_(account, calendars);
}

/**
 * Optional Web App HTTP trigger (same pattern as ProductivityTools.Spendings.*).
 * Deploy as Web App: Execute as = Me, Who has access = Only myself.
 */
function doGet() {
  var startedAt = Date.now();
  try {
    var summary = syncRawCalendarEvents();
    var seconds = Math.round((Date.now() - startedAt) / 1000);
    return ContentService
      .createTextOutput("OK: syncRawCalendarEvents finished in " + seconds + " s. " + JSON.stringify(summary))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    console.error("doGet failed:", err);
    return ContentService
      .createTextOutput("ERROR: " + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function rawSyncGetAccount_() {
  var account = Session.getEffectiveUser().getEmail();
  if (!account) {
    throw new Error("Could not resolve account email from Session.");
  }
  return account;
}

function rawSyncFetchPlan_(account) {
  var url = RAW_SYNC_API_URL + "/api/sync/plan?account=" + encodeURIComponent(account);
  var response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("GET /api/sync/plan failed (" + code + "): " + body);
  }

  var plan = JSON.parse(body);
  console.log("Sync plan for", account, "| calendars:", (plan.calendars || []).length);
  return plan;
}

function rawSyncExecuteCalendarPlans_(account, calendars) {
  var tz = Session.getScriptTimeZone();
  var calendarsById = {};
  for (var i = 0; i < calendars.length; i++) {
    calendarsById[calendars[i].calendarId] = calendars[i];
  }

  var pendingMutations = [];
  var pass1 = rawSyncSyncCalendarsOnce_(account, calendars, tz, pendingMutations);

  var mutationsApplied = 0;
  var mutationsFailed = 0;
  var calendarsResynced = 0;

  if (pendingMutations.length > 0) {
    var mutationOutcome = rawSyncApplyMutations_(pendingMutations);
    mutationsApplied = mutationOutcome.applied;
    mutationsFailed = mutationOutcome.failed;

    var resyncCalendars = [];
    for (var calId in mutationOutcome.affectedCalendarIds) {
      if (mutationOutcome.affectedCalendarIds.hasOwnProperty(calId) && calendarsById[calId]) {
        resyncCalendars.push(calendarsById[calId]);
      }
    }

    if (resyncCalendars.length > 0) {
      // Single follow-up pass (mutations ignored) so SQL immediately reflects moved/renamed events.
      var pass2 = rawSyncSyncCalendarsOnce_(account, resyncCalendars, tz, null);
      calendarsResynced = pass2.synced;
      pass1.failed += pass2.failed;
      pass1.totalEventsSent += pass2.totalEventsSent;
    }
  }

  var summary = {
    account: account,
    calendarsSynced: pass1.synced,
    calendarsFailed: pass1.failed,
    totalEventsSent: pass1.totalEventsSent,
    mutationsApplied: mutationsApplied,
    mutationsFailed: mutationsFailed,
    calendarsResynced: calendarsResynced
  };
  console.log("Raw calendar sync finished:", JSON.stringify(summary));
  return summary;
}

function rawSyncSyncCalendarsOnce_(account, calendars, tz, collectMutations) {
  var synced = 0;
  var failed = 0;
  var totalEventsSent = 0;

  for (var i = 0; i < calendars.length; i++) {
    var cal = calendars[i];
    var range = rawSyncFormatRange_(cal.startDate, cal.endDate, tz);

    try {
      var calendarName = rawSyncGetCalendarName_(cal.calendarId);
      var events = rawSyncFetchEvents_(cal.calendarId, range.startIso, range.endIso);

      var result = rawSyncPostBatch_({
        account: account,
        calendarId: cal.calendarId,
        calendarName: calendarName,
        rangeStart: range.startIso,
        rangeEnd: range.endIso,
        events: events
      });

      if (result) {
        synced++;
        totalEventsSent += events.length;
        if (collectMutations && result.mutations && result.mutations.length > 0) {
          for (var m = 0; m < result.mutations.length; m++) {
            collectMutations.push(result.mutations[m]);
          }
        }
      } else {
        failed++;
      }
    } catch (err) {
      failed++;
      console.error("Error syncing calendar", cal.calendarId, ":", err);
    }
  }

  return { synced: synced, failed: failed, totalEventsSent: totalEventsSent };
}

function rawSyncApplyMutations_(mutations) {
  var applied = 0;
  var failed = 0;
  var affectedCalendarIds = {};

  for (var i = 0; i < mutations.length; i++) {
    var cmd = mutations[i];
    try {
      if (cmd.action === "RENAME") {
        Calendar.Events.patch(
          { summary: cmd.newTitle },
          cmd.sourceCalendarId,
          cmd.eventId,
          { sendUpdates: "none" }
        );
        affectedCalendarIds[cmd.sourceCalendarId] = true;
        applied++;
      } else if (cmd.action === "MOVE") {
        rawSyncMoveEvent_(cmd);
        affectedCalendarIds[cmd.sourceCalendarId] = true;
        if (cmd.targetCalendarId) {
          affectedCalendarIds[cmd.targetCalendarId] = true;
        }
        applied++;
      } else {
        failed++;
        console.warn("Unknown mutation action:", JSON.stringify(cmd));
      }
    } catch (err) {
      failed++;
      console.error("Mutation failed:", JSON.stringify(cmd), "Error:", err);
    }
  }

  console.log("Mutations applied:", applied, "| failed:", failed);
  return { applied: applied, failed: failed, affectedCalendarIds: affectedCalendarIds };
}

function rawSyncMoveEvent_(cmd) {
  var moved = null;
  try {
    moved = Calendar.Events.move(
      cmd.sourceCalendarId,
      cmd.eventId,
      cmd.targetCalendarId,
      { sendUpdates: "none" }
    );
  } catch (moveErr) {
    console.warn("Calendar.Events.move failed for", cmd.eventId, "- falling back to copy+delete:", moveErr);
  }

  if (moved) {
    if (cmd.newTitle && moved.summary !== cmd.newTitle) {
      Calendar.Events.patch(
        { summary: cmd.newTitle },
        cmd.targetCalendarId,
        moved.id || cmd.eventId,
        { sendUpdates: "none" }
      );
    }
    return;
  }

  // Fallback for non-movable events (e.g. single instance of a recurring series):
  // insert into target first, and only delete from source after insert succeeds.
  var original = Calendar.Events.get(cmd.sourceCalendarId, cmd.eventId);
  var copy = {
    summary: cmd.newTitle || original.summary,
    start: original.start,
    end: original.end
  };
  if (original.description) copy.description = original.description;
  if (original.location) copy.location = original.location;
  if (original.colorId) copy.colorId = original.colorId;
  if (original.visibility) copy.visibility = original.visibility;
  if (original.transparency) copy.transparency = original.transparency;
  if (original.reminders) copy.reminders = original.reminders;
  if (original.attendees && original.attendees.length > 0) {
    copy.attendees = original.attendees
      .filter(function (a) { return a && a.email; })
      .map(function (a) {
        var guest = { email: a.email };
        if (a.optional) guest.optional = true;
        return guest;
      });
  }

  Calendar.Events.insert(copy, cmd.targetCalendarId, { sendUpdates: "none" });
  Calendar.Events.remove(cmd.sourceCalendarId, cmd.eventId, { sendUpdates: "none" });
}

function rawSyncFormatRange_(startDateStr, endDateStr, tz) {
  var s = startDateStr.split("-");
  var e = endDateStr.split("-");
  var start = new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]), 0, 0, 0, 0);
  var end = new Date(Number(e[0]), Number(e[1]) - 1, Number(e[2]), 23, 59, 59, 999);
  return {
    startIso: Utilities.formatDate(start, tz, "yyyy-MM-dd'T'HH:mm:ss.SSSXXX"),
    endIso: Utilities.formatDate(end, tz, "yyyy-MM-dd'T'HH:mm:ss.SSSXXX")
  };
}

function rawSyncGetCalendarName_(calendarId) {
  try {
    var entry = Calendar.CalendarList.get(calendarId);
    if (entry && (entry.summaryOverride || entry.summary)) {
      return entry.summaryOverride || entry.summary;
    }
  } catch (err) {
    // Fallback if calendar is not in CalendarList
  }

  var meta = Calendar.Calendars.get(calendarId);
  return (meta && meta.summary) ? meta.summary : calendarId;
}

function rawSyncFetchEvents_(calendarId, timeMin, timeMax) {
  var items = [];
  var pageToken = null;

  do {
    var response = Calendar.Events.list(calendarId, {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      maxResults: 2500,
      pageToken: pageToken
    });

    if (response.items && response.items.length > 0) {
      items = items.concat(response.items);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  return items;
}

function rawSyncPostBatch_(payload) {
  var response = UrlFetchApp.fetch(RAW_SYNC_API_URL + "/api/sync/raw-events", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code >= 200 && code < 300) {
    console.log("Synced:", payload.calendarName, "| sent:", payload.events.length, "| response:", body);
    try {
      return JSON.parse(body);
    } catch (err) {
      return {};
    }
  }

  console.error("POST /api/sync/raw-events failed for", payload.calendarId, "Status:", code, "Body:", body);
  return null;
}
