// Thin Microsoft Graph client for the pieces this app needs.

async function graphFetch(path, options) {
  options = options || {};
  var token = await getToken(CONFIG.scopes);
  var headers = Object.assign({
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json'
  }, options.headers || {});
  var res = await fetch(CONFIG.graphBase + path, {
    method: options.method || 'GET',
    headers: headers,
    body: options.body || undefined
  });
  if (!res.ok) {
    var msg = res.status + ' ' + res.statusText;
    try { var j = await res.json(); if (j && j.error && j.error.message) msg = j.error.message; } catch (e) {}
    throw new Error(msg);
  }
  if (res.status === 204) return null;
  var text = await res.text();
  return text ? JSON.parse(text) : null;
}

async function getMe() {
  return graphFetch('/me?$select=displayName,mail,userPrincipalName');
}

// Returns the signed-in agent's events for a window. Times come back as naive local
// strings in the requested time zone (via the Prefer header), which FullCalendar renders directly.
async function getCalendarView(startIso, endIso, timeZone) {
  var select = 'subject,start,end,showAs,isAllDay,location,onlineMeeting,organizer';
  var path = '/me/calendarView' +
    '?startDateTime=' + encodeURIComponent(startIso) +
    '&endDateTime=' + encodeURIComponent(endIso) +
    '&$select=' + select +
    '&$orderby=start/dateTime&$top=200';
  var data = await graphFetch(path, { headers: { 'Prefer': 'outlook.timezone="' + timeZone + '"' } });
  return data ? data.value : [];
}

// Creates an event on the agent's own calendar. Pass isOnlineMeeting/onlineMeetingProvider
// to have Exchange generate a Teams meeting; response carries onlineMeeting.joinUrl.
async function createTeamsEvent(evt) {
  return graphFetch('/me/events', { method: 'POST', body: JSON.stringify(evt) });
}

// Free/busy for one or more mailboxes across the org (delegated Calendars.Read, covered by
// Calendars.ReadWrite). Returns availability blocks only — no event details unless a mailbox
// has shared its calendar with the signed-in user.
async function getSchedule(emails, startDate, endDate) {
  var body = {
    schedules: emails,
    startTime: { dateTime: startDate.toISOString().slice(0, 19), timeZone: 'UTC' },
    endTime: { dateTime: endDate.toISOString().slice(0, 19), timeZone: 'UTC' },
    availabilityViewInterval: 30
  };
  var data = await graphFetch('/me/calendar/getSchedule', { method: 'POST', body: JSON.stringify(body) });
  return data ? data.value : [];
}
