// Jednorazowy skrypt migracji starych kalendarzy kategorii (np. Czerwiec 2025)
// do jednego kalendarza docelowego (np. DataPoints) ze zmianą tytułu na "<Prefiks>:<StaryTytuł>".
//
// Przykład:
//   Kalendarz źródłowy: "SelfDevelopment", wydarzenie: "Coding"
//   -> Przeniesione do kalendarza "DataPoints" z nowym tytułem: "SelfDevelopment:Coding"
//
// Jak uruchomić bezpiecznie w edytorze Apps Script:
//   1. (Opcjonalnie) Uruchom listAllMyCalendars() — wypisze w logu dokładne nazwy i ID wszystkich Twoich kalendarzy.
//   2. Sprawdź / dostosuj poniższy obiekt CONSOLIDATE_CONFIG.
//   3. Uruchom najpierw dryRunConsolidateCalendars() — nic nie zmienia w Google Calendar, tylko wypisuje w logu co zostanie przeniesione.
//   4. Gdy lista w logu się zgadza, uruchom executeConsolidateCalendars() — wykona właściwe przeniesienie i zmianę tytułów.

const CONSOLIDATE_CONFIG = {
  // Zakres dat (włącznie), z którego wydarzenia mają zostać przeniesione.
  // Dla samego czerwca 2025 możesz ustawić: startDate: "2025-06-01", endDate: "2025-06-30".
  startDate: "2025-01-01",
  endDate: "2025-12-31",

  // Domyślny kalendarz docelowy (nazwa kalendarza z Google Calendar LUB jego pełne CalendarId).
  defaultTargetCalendar: "DataPoints",

  // Separator między prefiksem a dotychczasowym tytułem wydarzenia (":" daje np. "SelfDevelopment:Coding").
  titleSeparator: ":",

  // Lista X kalendarzy źródłowych do przeiterowania.
  // Możesz podać nazwę kalendarza lub jego CalendarId (... @group.calendar.google.com).
  // Możesz dopisać dowolną liczbę kolejnych pozycji.
  rules: [
    {
      sourceCalendar: "SelfDevelopment",
      titlePrefix: "SelfDevelopment"
    },
    {
      sourceCalendar: "IndividualProjectWork",
      titlePrefix: "IndividualProjectWork"
    },
    {
      sourceCalendar: "Operations",
      titlePrefix: "Operations"
    }
  ]
};

/**
 * KROK 1 (Podgląd / Symulacja):
 * Wypisuje w Execution Log wszystkie wydarzenia, które zostaną przeniesione i ich nowe tytuły,
 * BEZ wprowadzania jakichkolwiek zmian w Google Calendar.
 */
function dryRunConsolidateCalendars() {
  return consolidateCalendarsRun_(true);
}

/**
 * KROK 2 (Właściwe wykonanie):
 * Przenosi wydarzenia ze źródłowych kalendarzy do kalendarza docelowego (np. DataPoints)
 * i zmienia ich tytuły na "<titlePrefix>:<staryTytuł>".
 */
function executeConsolidateCalendars() {
  return consolidateCalendarsRun_(false);
}

/**
 * Pomocnicza funkcja wypisująca wszystkie kalendarze na koncie (wraz z ukrytymi),
 * aby łatwo sprawdzić dokładną pisownię nazw lub skopiować CalendarId.
 */
function listAllMyCalendars() {
  var all = consolidateLoadAllCalendars_();
  console.log("Znaleziono kalendarzy:", all.length);
  for (var i = 0; i < all.length; i++) {
    var c = all[i];
    console.log(
      "[" + (i + 1) + "] Nazwa: \"" + c.name + "\" | ID: " + c.id + " | accessRole: " + c.accessRole
    );
  }
  return all;
}

function consolidateCalendarsRun_(dryRun) {
  var tz = Session.getScriptTimeZone();
  var range = consolidateFormatRange_(CONSOLIDATE_CONFIG.startDate, CONSOLIDATE_CONFIG.endDate, tz);
  var separator = CONSOLIDATE_CONFIG.titleSeparator !== undefined ? CONSOLIDATE_CONFIG.titleSeparator : ":";
  var allCalendars = consolidateLoadAllCalendars_();

  console.log(
    (dryRun ? "=== TRYB TESTOWY (DRY-RUN) ===" : "=== WYKONYWANIE MIGRACJI ===") +
    " | Zakres: " + CONSOLIDATE_CONFIG.startDate + " .. " + CONSOLIDATE_CONFIG.endDate +
    " (" + tz + ")"
  );

  var totalFound = 0;
  var totalMoved = 0;
  var totalFailed = 0;
  var perCalendarSummary = [];

  for (var r = 0; r < CONSOLIDATE_CONFIG.rules.length; r++) {
    var rule = CONSOLIDATE_CONFIG.rules[r];
    var targetSpec = rule.targetCalendar || CONSOLIDATE_CONFIG.defaultTargetCalendar;

    var sourceCal = consolidateResolveCalendar_(rule.sourceCalendar, allCalendars);
    if (!sourceCal) {
      console.error(
        "Pominięto regułę #" + (r + 1) + " — nie znaleziono kalendarza źródłowego: \"" +
        rule.sourceCalendar + "\". Dostępne kalendarze: " +
        allCalendars.map(function (c) { return "\"" + c.name + "\""; }).join(", ")
      );
      totalFailed++;
      continue;
    }

    var targetCal = consolidateResolveCalendar_(targetSpec, allCalendars);
    if (!targetCal) {
      throw new Error(
        "Nie znaleziono kalendarza docelowego: \"" + targetSpec + "\". " +
        "Uruchom listAllMyCalendars(), aby sprawdzić dokładną nazwę lub ID. Dostępne: " +
        allCalendars.map(function (c) { return "\"" + c.name + "\""; }).join(", ")
      );
    }

    if (sourceCal.id === targetCal.id) {
      console.warn("Kalendarz źródłowy i docelowy są identyczne (" + sourceCal.name + ") — pomijam.");
      continue;
    }

    var prefix = (rule.titlePrefix || sourceCal.name || "").trim();
    var events = consolidateFetchEvents_(sourceCal.id, range.startIso, range.endIso);

    console.log(
      "--- Kalendarz [" + (r + 1) + "/" + CONSOLIDATE_CONFIG.rules.length + "]: \"" +
      sourceCal.name + "\" -> \"" + targetCal.name + "\" | Znaleziono wydarzeń: " + events.length
    );

    var calMoved = 0;
    var calFailed = 0;

    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      if (!ev || ev.status === "cancelled") {
        continue;
      }

      totalFound++;
      var oldTitle = (ev.summary || "").trim();
      var newTitle = consolidateBuildNewTitle_(oldTitle, prefix, separator);
      var whenLabel = (ev.start && (ev.start.dateTime || ev.start.date)) || "?";

      if (dryRun) {
        console.log(
          "  [DRY-RUN] " + whenLabel + " | \"" + oldTitle + "\" => \"" + newTitle + "\"" +
          " (do: " + targetCal.name + ")"
        );
        calMoved++;
        totalMoved++;
        continue;
      }

      try {
        consolidateMoveAndRenameEvent_(sourceCal.id, targetCal.id, ev.id, newTitle);
        console.log(
          "  [OK] " + whenLabel + " | \"" + oldTitle + "\" => \"" + newTitle + "\""
        );
        calMoved++;
        totalMoved++;
        // Krótka pauza chroniąca przed limitem zapytań Google Calendar API przy dużej liczbie wydarzeń
        Utilities.sleep(80);
      } catch (err) {
        calFailed++;
        totalFailed++;
        console.error(
          "  [BŁĄD] Nie udało się przenieść wydarzenia id=" + ev.id +
          " (\"" + oldTitle + "\", " + whenLabel + "):",
          err
        );
      }
    }

    perCalendarSummary.push({
      source: sourceCal.name,
      target: targetCal.name,
      eventsProcessed: calMoved,
      eventsFailed: calFailed
    });
  }

  var summary = {
    dryRun: dryRun,
    startDate: CONSOLIDATE_CONFIG.startDate,
    endDate: CONSOLIDATE_CONFIG.endDate,
    totalFound: totalFound,
    totalProcessed: totalMoved,
    totalFailed: totalFailed,
    calendars: perCalendarSummary
  };

  console.log("=== PODSUMOWANIE ===", JSON.stringify(summary, null, 2));
  return summary;
}

function consolidateBuildNewTitle_(oldTitle, prefix, separator) {
  if (!oldTitle) {
    return prefix;
  }
  // Zabezpieczenie przed podwójnym dodaniem prefiksu (np. gdyby tytuł miał już "SelfDevelopment:...")
  var normalizedLower = oldTitle.toLowerCase();
  var checkColon = (prefix + ":").toLowerCase();
  if (normalizedLower.indexOf(checkColon) === 0) {
    return oldTitle;
  }
  return prefix + separator + oldTitle;
}

function consolidateMoveAndRenameEvent_(sourceCalendarId, targetCalendarId, eventId, newTitle) {
  var moved = null;
  try {
    moved = Calendar.Events.move(
      sourceCalendarId,
      eventId,
      targetCalendarId,
      { sendUpdates: "none" }
    );
  } catch (moveErr) {
    console.warn(
      "Calendar.Events.move nie powiodło się dla eventId=" + eventId +
      " (np. pojedyncze wystąpienie wydarzenia cyklicznego) — używam bezpiecznego fallbacku insert+remove:",
      moveErr
    );
  }

  if (moved) {
    if (moved.summary !== newTitle) {
      Calendar.Events.patch(
        { summary: newTitle },
        targetCalendarId,
        moved.id || eventId,
        { sendUpdates: "none" }
      );
    }
    return;
  }

  // Fallback dla wydarzeń, których API nie pozwala przenieść przez Events.move:
  // najpierw tworzymy kopię w kalendarzu docelowym, a dopiero po sukcesie usuwamy oryginał.
  var original = Calendar.Events.get(sourceCalendarId, eventId);
  var copy = {
    summary: newTitle,
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

  Calendar.Events.insert(copy, targetCalendarId, { sendUpdates: "none" });
  Calendar.Events.remove(sourceCalendarId, eventId, { sendUpdates: "none" });
}

function consolidateLoadAllCalendars_() {
  var result = [];
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
      result.push({
        id: item.id,
        name: item.summaryOverride || item.summary || item.id,
        summary: item.summary || "",
        summaryOverride: item.summaryOverride || "",
        accessRole: item.accessRole || ""
      });
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);

  return result;
}

function consolidateResolveCalendar_(nameOrId, allCalendars) {
  if (!nameOrId) return null;
  var needle = nameOrId.toString().trim();
  var needleLower = needle.toLowerCase();

  // 1. Dokładne dopasowanie po CalendarId
  for (var i = 0; i < allCalendars.length; i++) {
    if (allCalendars[i].id === needle) {
      return allCalendars[i];
    }
  }

  // 2. Dopasowanie po nazwie (bez rozróżniania wielkości liter)
  for (var j = 0; j < allCalendars.length; j++) {
    var c = allCalendars[j];
    if (
      c.name.toLowerCase() === needleLower ||
      c.summary.toLowerCase() === needleLower ||
      c.summaryOverride.toLowerCase() === needleLower
    ) {
      return c;
    }
  }

  return null;
}

function consolidateFetchEvents_(calendarId, timeMin, timeMax) {
  var items = [];
  var pageToken = null;

  do {
    var response = Calendar.Events.list(calendarId, {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      showDeleted: false,
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

function consolidateFormatRange_(startDateStr, endDateStr, tz) {
  var s = startDateStr.split("-");
  var e = endDateStr.split("-");
  var start = new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]), 0, 0, 0, 0);
  var end = new Date(Number(e[0]), Number(e[1]) - 1, Number(e[2]), 23, 59, 59, 999);
  return {
    startIso: Utilities.formatDate(start, tz, "yyyy-MM-dd'T'HH:mm:ss.SSSXXX"),
    endIso: Utilities.formatDate(end, tz, "yyyy-MM-dd'T'HH:mm:ss.SSSXXX")
  };
}
