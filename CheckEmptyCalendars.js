// Skrypt weryfikujący (TYLKO DO ODCZYTU), czy stare kalendarze są całkowicie puste
// w całej historii i przyszłości (bez żadnych ograniczeń dat) przed ich ręcznym usunięciem.
//
// Jak użyć w edytorze Google Apps Script:
//   1. Wpisz poniżej w CALENDARS_TO_CHECK nazwy lub CalendarId kalendarzy, które chcesz usunąć.
//   2. Wybierz z listy na górze funkcję: checkCalendarsBeforeDelete i kliknij "Uruchom".
//   3. Sprawdź Execution Log — dla każdego kalendarza zobaczysz status:
//      - [PUSTY - BEZPIECZNY DO USUNIĘCIA] (0 wydarzeń w całej historii)
//      - [UWAGA - ZAWIERA WYDARZENIA] wraz z listą dat i tytułów pozostałych eventów.

const CALENDARS_TO_CHECK = [
  "SelfDevelopment",
  "IndividualProjectWork",
  "Operations",
  "TopDown",
  "FocusedWork",
  "DataPoints"
];

// Ile maksymalnie przykładowych wydarzeń wypisać w logu dla niepustego kalendarza
const MAX_SAMPLE_EVENTS_TO_LOG = 30;

/**
 * Główna funkcja sprawdzająca wszystkie kalendarze z listy CALENDARS_TO_CHECK
 * w całym zakresie czasu (przeszłość + przyszłość). Nic nie modyfikuje ani nie usuwa.
 */
function checkCalendarsBeforeDelete() {
  var allCalendars = checkEmptyLoadAllCalendars_();
  var safeToDelete = [];
  var notEmpty = [];
  var notFound = [];

  console.log("=== SPRAWDZANIE KALENDARZY PRZED USUNIĘCIEM (CAŁA HISTORIA + PRZYSZŁOŚĆ) ===");

  for (var i = 0; i < CALENDARS_TO_CHECK.length; i++) {
    var spec = CALENDARS_TO_CHECK[i];
    var cal = checkEmptyResolveCalendar_(spec, allCalendars);

    if (!cal) {
      console.warn("❓ NIE ZNALEZIONO KALENDARZA: \"" + spec + "\" (być może został już usunięty lub ma inną nazwę).");
      notFound.push(spec);
      continue;
    }

    var activeEvents = checkEmptyFetchAllActiveEvents_(cal.id);

    if (activeEvents.length === 0) {
      console.log("✅ PUSTY — BEZPIECZNY DO USUNIĘCIA: \"" + cal.name + "\" (ID: " + cal.id + ")");
      safeToDelete.push({ name: cal.name, id: cal.id });
    } else {
      var dateBounds = checkEmptyComputeDateBounds_(activeEvents);
      console.warn(
        "⚠️ UWAGA — NIE USUWAJ! Kalendarz \"" + cal.name + "\" zawiera jeszcze " +
        activeEvents.length + " wydarzeń (zakres dat: " + dateBounds.minDate + " .. " + dateBounds.maxDate + "):"
      );

      var limit = Math.min(activeEvents.length, MAX_SAMPLE_EVENTS_TO_LOG);
      for (var e = 0; e < limit; e++) {
        var ev = activeEvents[e];
        console.log(
          "   [" + (e + 1) + "] " + ev.whenLabel +
          " | \"" + ev.title + "\"" +
          (ev.isRecurring ? " (cykliczne)" : "") +
          " | id=" + ev.id
        );
      }
      if (activeEvents.length > limit) {
        console.log("   ... oraz " + (activeEvents.length - limit) + " kolejnych wydarzeń.");
      }

      notEmpty.push({
        name: cal.name,
        id: cal.id,
        remainingEventsCount: activeEvents.length,
        earliestEvent: dateBounds.minDate,
        latestEvent: dateBounds.maxDate
      });
    }
  }

  var summary = {
    safeToDeleteCount: safeToDelete.length,
    safeToDelete: safeToDelete.map(function (c) { return c.name; }),
    notEmptyCount: notEmpty.length,
    notEmpty: notEmpty,
    notFound: notFound
  };

  console.log("=== PODSUMOWANIE WERYFIKACJI ===", JSON.stringify(summary, null, 2));
  return summary;
}

function checkEmptyFetchAllActiveEvents_(calendarId) {
  var active = [];
  var pageToken = null;

  do {
    // Brak timeMin/timeMax oraz singleEvents=false pobiera WSZYSTKIE nieusunięte wpisy z całej historii i przyszłości
    var response = Calendar.Events.list(calendarId, {
      showDeleted: false,
      maxResults: 2500,
      pageToken: pageToken
    });

    var items = response.items || [];
    for (var i = 0; i < items.length; i++) {
      var ev = items[i];
      if (!ev || ev.status === "cancelled") {
        continue;
      }

      var isRecurring = Boolean(ev.recurrence && ev.recurrence.length > 0);

      // Jeśli to główny wpis serii cyklicznej, sprawdzamy czy ma jeszcze jakiekolwiek nieusunięte wystąpienia
      if (isRecurring) {
        var instResp = Calendar.Events.instances(calendarId, ev.id, {
          showDeleted: false,
          maxResults: 1
        });
        if (!instResp.items || instResp.items.length === 0) {
          continue;
        }
      }

      var whenLabel =
        (ev.start && (ev.start.dateTime || ev.start.date)) || "(brak daty startu)";
      var dayOnly = whenLabel.length >= 10 ? whenLabel.substring(0, 10) : whenLabel;

      active.push({
        id: ev.id,
        title: (ev.summary || "(bez tytułu)").trim(),
        whenLabel: whenLabel,
        dayOnly: dayOnly,
        isRecurring: isRecurring
      });
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  active.sort(function (a, b) {
    return a.whenLabel.localeCompare(b.whenLabel);
  });

  return active;
}

function checkEmptyComputeDateBounds_(events) {
  if (!events || events.length === 0) {
    return { minDate: "-", maxDate: "-" };
  }
  var days = events
    .map(function (e) { return e.dayOnly; })
    .filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); })
    .sort();

  if (days.length === 0) {
    return { minDate: "?", maxDate: "?" };
  }
  return {
    minDate: days[0],
    maxDate: days[days.length - 1]
  };
}

function checkEmptyLoadAllCalendars_() {
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
        summaryOverride: item.summaryOverride || ""
      });
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);

  return result;
}

function checkEmptyResolveCalendar_(nameOrId, allCalendars) {
  if (!nameOrId) return null;
  var needle = nameOrId.toString().trim();
  var needleLower = needle.toLowerCase();

  for (var i = 0; i < allCalendars.length; i++) {
    if (allCalendars[i].id === needle) {
      return allCalendars[i];
    }
  }

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
