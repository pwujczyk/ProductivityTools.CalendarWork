// Jednorazowy skrypt przemianowujący dotychczasowe wydarzenia w kalendarzu "DailyLog"
// z nazw sklejonych / skrótów na nazwy ze spacjami:
//   - "SelfDevelopment:..." / "SD:..."        => "Self Development: ..."
//   - "IndividualProjectWork:..." / "IPW:..." => "Individual Project Work: ..."
//   - "FocusedWork:..." / "FW:..."            => "Focused Work: ..."
//
// Jak użyć w edytorze Google Apps Script:
//   1. Uruchom dryRunRenameDailyLogToSpaced() — wypisze w logu podgląd zmian bez modyfikacji Google Calendar.
//   2. Uruchom executeRenameDailyLogToSpaced() — zaktualizuje tytuły w Google Calendar.

const RENAME_SPACED_CONFIG = {
  calendarNameOrId: "DailyLog",
  startDate: "2024-01-01",
  endDate: "2026-12-31",

  // Mapowanie prefiksów przed dwukropkiem (lub całych tytułów bez dwukropka) -> docelowy prefiks ze spacjami
  prefixMap: {
    "selfdevelopment": "Self Development",
    "sd": "Self Development",
    "individualprojectwork": "Individual Project Work",
    "ipw": "Individual Project Work",
    "focusedwork": "Focused Work",
    "fw": "Focused Work"
  }
};

function dryRunRenameDailyLogToSpaced() {
  return renameDailyLogToSpacedRun_(true);
}

function executeRenameDailyLogToSpaced() {
  return renameDailyLogToSpacedRun_(false);
}

function renameDailyLogToSpacedRun_(dryRun) {
  var tz = Session.getScriptTimeZone();
  var s = RENAME_SPACED_CONFIG.startDate.split("-");
  var e = RENAME_SPACED_CONFIG.endDate.split("-");
  var timeMin = Utilities.formatDate(
    new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]), 0, 0, 0, 0),
    tz,
    "yyyy-MM-dd'T'HH:mm:ss.SSSXXX"
  );
  var timeMax = Utilities.formatDate(
    new Date(Number(e[0]), Number(e[1]) - 1, Number(e[2]), 23, 59, 59, 999),
    tz,
    "yyyy-MM-dd'T'HH:mm:ss.SSSXXX"
  );

  var cal = renameSpacedResolveCalendar_(RENAME_SPACED_CONFIG.calendarNameOrId);
  if (!cal) {
    throw new Error("Nie znaleziono kalendarza: " + RENAME_SPACED_CONFIG.calendarNameOrId);
  }

  var events = [];
  var pageToken = null;
  do {
    var resp = Calendar.Events.list(cal.id, {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      showDeleted: false,
      maxResults: 2500,
      pageToken: pageToken
    });
    if (resp.items && resp.items.length > 0) {
      events = events.concat(resp.items);
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);

  console.log(
    (dryRun ? "=== DRY-RUN ===" : "=== EXECUTE ===") +
    " Kalendarz: \"" + cal.name + "\" | Pobranie wydarzeń: " + events.length
  );

  var matched = 0;
  var updated = 0;
  var failed = 0;

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (!ev || ev.status === "cancelled" || !ev.summary) {
      continue;
    }

    var oldTitle = ev.summary.trim();
    var colonIdx = oldTitle.indexOf(":");
    var rawPrefix = colonIdx === -1 ? oldTitle : oldTitle.substring(0, colonIdx).trim();
    var key = rawPrefix.toLowerCase();

    if (!RENAME_SPACED_CONFIG.prefixMap.hasOwnProperty(key)) {
      continue;
    }

    var targetPrefix = RENAME_SPACED_CONFIG.prefixMap[key];
    var newTitle;
    if (colonIdx === -1) {
      newTitle = targetPrefix;
    } else {
      var suffix = oldTitle.substring(colonIdx + 1).trim();
      newTitle = suffix ? (targetPrefix + ": " + suffix) : targetPrefix;
    }

    if (newTitle === ev.summary) {
      continue;
    }

    matched++;
    var whenLabel = (ev.start && (ev.start.dateTime || ev.start.date)) || "?";

    if (dryRun) {
      console.log("  [DRY-RUN] " + whenLabel + " | \"" + ev.summary + "\" => \"" + newTitle + "\"");
      continue;
    }

    try {
      Calendar.Events.patch(
        { summary: newTitle },
        cal.id,
        ev.id,
        { sendUpdates: "none" }
      );
      updated++;
      if (updated % 25 === 0) {
        console.log("  Zaktualizowano dotychczas: " + updated + " / " + matched + "...");
      }
      Utilities.sleep(60);
    } catch (err) {
      failed++;
      console.error("  [BŁĄD] id=" + ev.id + " (\"" + ev.summary + "\"):", err);
    }
  }

  var summary = {
    dryRun: dryRun,
    calendar: cal.name,
    eventsChecked: events.length,
    eventsToRename: matched,
    eventsUpdated: dryRun ? 0 : updated,
    eventsFailed: failed
  };
  console.log("=== PODSUMOWANIE ===", JSON.stringify(summary, null, 2));
  return summary;
}

function renameSpacedResolveCalendar_(nameOrId) {
  var needle = nameOrId.trim();
  var needleLower = needle.toLowerCase();
  var pageToken = null;

  do {
    var resp = Calendar.CalendarList.list({
      showHidden: true,
      maxResults: 250,
      pageToken: pageToken
    });
    var items = resp.items || [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var name = item.summaryOverride || item.summary || item.id;
      if (
        item.id === needle ||
        name.toLowerCase() === needleLower ||
        (item.summary || "").toLowerCase() === needleLower
      ) {
        return { id: item.id, name: name };
      }
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);

  return null;
}
