(function () {
  'use strict';

  var calendar = null;
  var selected = [];                // emails currently shown; empty = my calendar
  var MEMBERS = [];                 // [{ name, email, color }] built from CONFIG.teamMembers
  var PALETTE = ['#2f5bea', '#c4314b', '#3a9b78', '#8764b8', '#b8860b', '#0f766e', '#b5417a', '#3b6ea5'];
  var els = {};

  var SVG_VIDEO = '<svg class="i i-sm" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M1 5.5A2.5 2.5 0 0 1 3.5 3h4A2.5 2.5 0 0 1 10 5.5v5A2.5 2.5 0 0 1 7.5 13h-4A2.5 2.5 0 0 1 1 10.5zm12.04 6.28L10.9 10.3l.02-.28V5.98q0-.15-.02-.28l2.14-1.48c.82-.57 1.96.02 1.96 1.03v5.5c0 1-1.14 1.6-1.96 1.03"/></svg>';

  function $(id) { return document.getElementById(id); }
  function show(el) { if (el) el.classList.remove('hidden'); }
  function hide(el) { if (el) el.classList.add('hidden'); }

  document.addEventListener('DOMContentLoaded', function () {
    cacheEls();
    wireEvents();
    if (CONFIG.allowViewOthers) buildPicker(); else hide(els.subbar);
    boot();
  });

  function cacheEls() {
    ['statusBar', 'statusText', 'signInBtn', 'who', 'refreshBtn', 'newBtn', 'calendar',
     'subbar', 'ctxLabel', 'selChips', 'pickerBtn', 'pickerMenu', 'pickerList', 'pickMine',
     'bookingModal', 'bookingForm', 'bmSubject', 'bmStart', 'bmEnd', 'bmAttendee',
     'bmLocation', 'bmNotes', 'bmTeams', 'bmCancel', 'bmSave', 'bmError',
     'detailsModal', 'dmTitle', 'dmBody', 'dmClose', 'toast']
      .forEach(function (id) { els[id] = $(id); });
  }

  function wireEvents() {
    els.signInBtn.addEventListener('click', onSignInClick);
    els.refreshBtn.addEventListener('click', function () { if (calendar) calendar.refetchEvents(); });
    els.newBtn.addEventListener('click', function () { openBooking(defaultStart(), defaultEnd()); });
    els.bmCancel.addEventListener('click', function () { closeModal(els.bookingModal); });
    els.bookingForm.addEventListener('submit', onBookingSubmit);
    els.dmClose.addEventListener('click', function () { closeModal(els.detailsModal); });
    [els.bookingModal, els.detailsModal].forEach(function (m) {
      m.addEventListener('click', function (e) { if (e.target === m) closeModal(m); });
    });
    window.addEventListener('resize', function () { if (calendar) calendar.updateSize(); });
    window.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') { closeModal(els.bookingModal); closeModal(els.detailsModal); }
    });
    els.pickerBtn.addEventListener('click', function (e) { e.stopPropagation(); togglePicker(); });
    els.pickMine.addEventListener('click', selectMine);
    document.addEventListener('click', function (e) { if (!e.target.closest('.picker')) closePicker(); });
  }

  async function boot() {
    if (!isConfigured()) {
      setStatus('This app is not configured yet. Set clientId and tenantId in config.js.', true);
      return;
    }
    initAuth();
    try {
      setStatus('Signing you in', false, true);
      await ensureSignedInSilent();
      await afterSignIn();
    } catch (e) {
      setStatus('Please sign in to view your Outlook calendar.', false);
      show(els.signInBtn);
    }
  }

  async function onSignInClick() {
    hide(els.signInBtn);
    try {
      setStatus('Signing you in', false, true);
      await interactiveSignIn();
      await afterSignIn();
    } catch (e) {
      setStatus('Sign-in failed: ' + friendlyError(e), true);
      show(els.signInBtn);
    }
  }

  async function afterSignIn() {
    clearStatus();
    try {
      var me = await getMe();
      els.who.textContent = me.displayName || me.userPrincipalName || 'Signed in';
      els.who.title = me.mail || me.userPrincipalName || '';
    } catch (e) { /* header name is non-essential */ }
    if (!calendar) initCalendar(); else calendar.refetchEvents();
  }

  function initCalendar() {
    var bh = CONFIG.businessHours || { start: '08:00', end: '18:00' };
    calendar = new FullCalendar.Calendar(els.calendar, {
      initialView: 'timeGridWeek',
      height: '100%',
      nowIndicator: true,
      firstDay: 1,
      slotMinTime: '07:00:00',
      slotMaxTime: '20:00:00',
      scrollTime: '08:00:00',
      allDaySlot: true,
      expandRows: true,
      stickyHeaderDates: true,
      headerToolbar: { left: 'prev,next today', center: 'title', right: 'timeGridDay,timeGridWeek,dayGridMonth' },
      businessHours: { daysOfWeek: [1, 2, 3, 4, 5], startTime: bh.start, endTime: bh.end },
      selectable: true,
      selectMirror: true,
      select: function (info) { openBooking(info.start, info.end); calendar.unselect(); },
      eventClick: onEventClick,
      events: fetchEvents,
      eventDidMount: function (arg) {
        var ep = arg.event.extendedProps;
        if (!ep.availabilityOnly) arg.el.classList.add('show-' + (ep.showAs || 'busy'));
        if (ep.showAs === 'tentative') arg.el.classList.add('striped');
        if (ep.online) {
          arg.el.classList.add('is-online');
          var t = arg.el.querySelector('.fc-event-title');
          if (t && !t.querySelector('.i')) t.insertAdjacentHTML('afterbegin', SVG_VIDEO + ' ');
        }
      }
    });
    calendar.render();
  }

  async function fetchEvents(info, success, failure) {
    try {
      if (selected.length) {
        var sched = CONFIG.demoAvailability
          ? selected.map(function (email) { return mockScheduleFor(email, info.start, info.end); })
          : await getSchedule(selected, info.start, info.end);
        var out = [];
        sched.forEach(function (s, i) {
          var m = memberFor(selected[i]);
          (s.scheduleItems || []).forEach(function (si) {
            var status = si.status || 'busy';
            if (status === 'free') return;
            out.push({
              title: m.name + (status !== 'busy' ? ' \u00b7 ' + statusLabel(status) : ''),
              start: schedDate(si.start),
              end: schedDate(si.end),
              backgroundColor: m.color,
              borderColor: m.color,
              textColor: '#fff',
              extendedProps: { showAs: status, person: m.name, availabilityOnly: true }
            });
          });
        });
        success(out);
      } else {
        var tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        var evts = await getCalendarView(info.start.toISOString(), info.end.toISOString(), tz);
        success((evts || []).map(mapEvent));
      }
    } catch (e) {
      setStatus('Could not load calendar: ' + friendlyError(e), true);
      failure(e);
    }
  }

  // --- Calendar picker (manager view) ---
  function buildMembers() {
    MEMBERS = (CONFIG.teamMembers || []).map(function (m, i) {
      return { name: m.name, email: m.email, color: PALETTE[i % PALETTE.length] };
    });
  }
  function memberFor(email) {
    for (var i = 0; i < MEMBERS.length; i++) { if (MEMBERS[i].email === email) return MEMBERS[i]; }
    return { name: email, email: email, color: PALETTE[0] };
  }
  function buildPicker() {
    buildMembers();
    els.pickerList.textContent = '';
    MEMBERS.forEach(function (m) {
      var row = document.createElement('label');
      row.className = 'picker-row';
      var cb = document.createElement('input');
      cb.type = 'checkbox'; cb.value = m.email;
      cb.addEventListener('change', onMemberToggle);
      var dot = document.createElement('span'); dot.className = 'dot'; dot.style.background = m.color;
      var txt = document.createElement('span'); txt.className = 'pr-text';
      var nm = document.createElement('span'); nm.className = 'pr-name'; nm.textContent = m.name;
      var em = document.createElement('span'); em.className = 'pr-email'; em.textContent = m.email;
      txt.appendChild(nm); txt.appendChild(em);
      row.appendChild(cb); row.appendChild(dot); row.appendChild(txt);
      els.pickerList.appendChild(row);
    });
  }
  function onMemberToggle(e) {
    var email = e.target.value;
    if (e.target.checked) { if (selected.indexOf(email) < 0) selected.push(email); }
    else { selected = selected.filter(function (x) { return x !== email; }); }
    afterSelectionChange();
  }
  function selectMine() {
    selected = [];
    Array.prototype.forEach.call(els.pickerList.querySelectorAll('input[type=checkbox]'), function (cb) { cb.checked = false; });
    closePicker();
    afterSelectionChange();
  }
  function removeSelected(email) {
    selected = selected.filter(function (x) { return x !== email; });
    var cb = els.pickerList.querySelector('input[value="' + email + '"]');
    if (cb) cb.checked = false;
    afterSelectionChange();
  }
  function afterSelectionChange() {
    renderChips();
    updateCtx();
    if (calendar) calendar.refetchEvents();
  }
  function renderChips() {
    els.selChips.textContent = '';
    selected.forEach(function (email) {
      var m = memberFor(email);
      var chip = document.createElement('span'); chip.className = 'chip';
      var dot = document.createElement('span'); dot.className = 'dot'; dot.style.background = m.color;
      var nm = document.createElement('span'); nm.textContent = m.name;
      var x = document.createElement('button'); x.type = 'button'; x.textContent = '\u00d7'; x.title = 'Remove ' + m.name;
      x.addEventListener('click', function () { removeSelected(email); });
      chip.appendChild(dot); chip.appendChild(nm); chip.appendChild(x);
      els.selChips.appendChild(chip);
    });
  }
  function updateCtx() {
    if (!selected.length) els.ctxLabel.textContent = 'My calendar';
    else if (selected.length === 1) els.ctxLabel.textContent = 'Viewing: ' + memberFor(selected[0]).name;
    else els.ctxLabel.textContent = 'Viewing ' + selected.length + ' calendars';
  }
  function togglePicker() { els.pickerMenu.classList.toggle('hidden'); }
  function closePicker() { els.pickerMenu.classList.add('hidden'); }

  // getSchedule returns availability blocks (status only), returned in UTC.
  function schedDate(dtz) {
    var s = dtz.dateTime, t = s.length > 11 ? s.slice(11) : '';
    return /[Z+\-]/.test(t) ? new Date(s) : new Date(s + 'Z');
  }
  function statusLabel(s) {
    return ({ busy: 'Busy', tentative: 'Tentative', oof: 'Out of office', workingElsewhere: 'Working elsewhere', free: 'Free' })[s] || 'Busy';
  }

  // --- Demo availability (fictitious roster) ---
  // Deterministic per-person busy/tentative/OOF blocks so the manager multi-calendar view is
  // populated without real Outlook data. Returns the same shape as getSchedule (UTC ISO + status).
  function hashCode(str) {
    var h = 0;
    for (var i = 0; i < str.length; i++) { h = (Math.imul(h, 31) + str.charCodeAt(i)) | 0; }
    return Math.abs(h);
  }
  function seededRand(seed) {
    return function () {
      seed = (seed + 0x6D2B79F5) | 0;
      var t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  function pushMockItem(items, dayDate, sh, sm, durMin, status) {
    var s = new Date(dayDate); s.setHours(sh, sm, 0, 0);
    var e = new Date(s.getTime() + durMin * 60000);
    items.push({ status: status, start: { dateTime: s.toISOString() }, end: { dateTime: e.toISOString() } });
  }
  function mockScheduleFor(email, rangeStart, rangeEnd) {
    var rnd = seededRand(hashCode(email) + 7);
    var density = 0.55 + rnd() * 0.35;                 // per-person busyness
    // [hour, minute, durationMin, baseProbability]
    var slots = [[8, 30, 30, 0.35], [9, 0, 15, 0.85], [9, 30, 60, 0.5], [11, 0, 60, 0.5],
                 [13, 0, 30, 0.4], [14, 0, 60, 0.55], [15, 30, 45, 0.4], [16, 30, 30, 0.35]];
    var items = [];
    var d = new Date(rangeStart); d.setHours(0, 0, 0, 0);
    var guard = 0;
    while (d < rangeEnd && guard++ < 60) {
      var dow = d.getDay();
      if (dow >= 1 && dow <= 5) {
        if (rnd() < 0.06) {
          pushMockItem(items, d, 9, 0, 8 * 60, 'oof');   // out-of-office day
        } else {
          slots.forEach(function (sl) {
            if (rnd() < sl[3] * density) {
              var r = rnd(), st = 'busy';
              if (r < 0.18) st = 'tentative';
              else if (r < 0.24) st = 'workingElsewhere';
              pushMockItem(items, d, sl[0], sl[1], sl[2], st);
            }
          });
        }
      }
      d.setDate(d.getDate() + 1);
    }
    return { scheduleItems: items };
  }

  function trimDt(s) { return s ? s.replace(/\.\d+$/, '') : s; }

  function mapEvent(e) {
    return {
      id: e.id,
      title: e.subject || '(No subject)',
      start: trimDt(e.start.dateTime),
      end: trimDt(e.end.dateTime),
      allDay: !!e.isAllDay,
      extendedProps: {
        showAs: e.showAs || 'busy',
        online: !!e.onlineMeeting,
        joinUrl: e.onlineMeeting ? e.onlineMeeting.joinUrl : null,
        location: e.location ? e.location.displayName : '',
        organizer: (e.organizer && e.organizer.emailAddress) ? e.organizer.emailAddress.name : ''
      }
    };
  }

  function onEventClick(arg) {
    var e = arg.event, p = e.extendedProps;
    els.dmTitle.textContent = e.title;
    var dl = els.dmBody;
    dl.textContent = '';
    addDetail(dl, 'When', formatRange(e.start, e.end, e.allDay));
    addDetail(dl, 'Status', capitalize(p.showAs));
    if (p.location) addDetail(dl, 'Location', p.location);
    if (p.organizer) addDetail(dl, 'Organizer', p.organizer);
    if (p.joinUrl && /^https:\/\//i.test(p.joinUrl)) {
      var dt = document.createElement('dt'); dt.textContent = 'Teams';
      var dd = document.createElement('dd');
      var a = document.createElement('a');
      a.href = p.joinUrl; a.target = '_blank'; a.rel = 'noopener';
      a.textContent = 'Join Teams meeting';
      a.insertAdjacentHTML('afterbegin', SVG_VIDEO + ' ');
      dd.appendChild(a); dl.appendChild(dt); dl.appendChild(dd);
    }
    openModal(els.detailsModal);
  }

  function addDetail(dl, label, value) {
    var dt = document.createElement('dt'); dt.textContent = label;
    var dd = document.createElement('dd'); dd.textContent = value;
    dl.appendChild(dt); dl.appendChild(dd);
  }

  function openBooking(start, end) {
    hide(els.bmError); els.bmError.textContent = '';
    els.bookingForm.reset();
    els.bmTeams.checked = true;
    els.bmStart.value = toLocalInput(start);
    els.bmEnd.value = toLocalInput(end || new Date(start.getTime() + (CONFIG.defaultMeetingMinutes || 30) * 60000));
    els.bmAttendee.value = prefillAttendee() || selected.join('; ');
    openModal(els.bookingModal);
    setTimeout(function () { els.bmSubject.focus(); }, 30);
  }

  // Optional: form-embed scenarios can pass a customer email via ?email= or ?attendee=.
  function prefillAttendee() {
    try {
      var p = new URLSearchParams(location.search);
      return p.get('email') || p.get('attendee') || '';
    } catch (e) { return ''; }
  }

  // Optional: link the Dynamics appointment to a record via ?regardingtype=incident&regardingid=<guid>.
  function regardingFromUrl() {
    try {
      var p = new URLSearchParams(location.search);
      var t = p.get('regardingtype'), id = p.get('regardingid');
      return (t && id) ? { entity: t, id: id } : null;
    } catch (e) { return null; }
  }

  async function onBookingSubmit(ev) {
    ev.preventDefault();
    hide(els.bmError); els.bmError.textContent = '';

    var subject = els.bmSubject.value.trim();
    if (!subject) return formError('Please enter a subject.');
    var start = new Date(els.bmStart.value), end = new Date(els.bmEnd.value);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return formError('Please provide valid start and end times.');
    if (end <= start) return formError('End time must be after the start time.');

    var tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    var payload = {
      subject: subject,
      body: { contentType: 'text', content: els.bmNotes.value || '' },
      start: { dateTime: graphDt(els.bmStart.value), timeZone: tz },
      end: { dateTime: graphDt(els.bmEnd.value), timeZone: tz }
    };
    var loc = els.bmLocation.value.trim();
    if (loc) payload.location = { displayName: loc };
    var attendee = els.bmAttendee.value.trim();
    if (attendee) {
      payload.attendees = attendee.split(/[;,]/).map(function (a) { return a.trim(); })
        .filter(Boolean).map(function (a) { return { emailAddress: { address: a }, type: 'required' }; });
    }
    if (els.bmTeams.checked) { payload.isOnlineMeeting = true; payload.onlineMeetingProvider = 'teamsForBusiness'; }

    els.bmSave.disabled = true; els.bmSave.textContent = 'Creating';
    try {
      var created = await createTeamsEvent(payload);
      var join = (created && created.onlineMeeting) ? created.onlineMeeting.joinUrl : null;

      var extraNote = '';
      if (CONFIG.createDynamicsAppointment && dataverseAvailable()) {
        try {
          var reg = regardingFromUrl();
          await createDynamicsAppointment({
            subject: subject,
            start: start,
            end: end,
            notes: els.bmNotes.value || '',
            location: loc,
            joinUrl: join,
            attendeeEmail: attendee,
            regardingEntity: reg ? reg.entity : null,
            regardingId: reg ? reg.id : null
          });
        } catch (dvErr) {
          extraNote = ' Dynamics appointment was not created: ' + friendlyError(dvErr);
        }
      }

      closeModal(els.bookingModal);
      showToast('Appointment created' + (join ? ' with a Teams meeting.' : '.') + extraNote, join);
      if (calendar) calendar.refetchEvents();
    } catch (e) {
      formError(friendlyError(e));
    } finally {
      els.bmSave.disabled = false; els.bmSave.textContent = 'Create appointment';
    }
  }

  function formError(msg) { els.bmError.textContent = msg; show(els.bmError); return false; }

  // datetime-local value ("2026-08-19T09:00") -> Graph dateTime ("2026-08-19T09:00:00").
  function graphDt(v) { return (v && v.length === 16) ? v + ':00' : v; }

  function toLocalInput(d) {
    var p = function (n) { return String(n).padStart(2, '0'); };
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
      'T' + p(d.getHours()) + ':' + p(d.getMinutes());
  }

  function defaultStart() {
    var d = new Date(); d.setSeconds(0, 0);
    d.setMinutes(d.getMinutes() < 30 ? 30 : 60);
    return d;
  }
  function defaultEnd() {
    return new Date(defaultStart().getTime() + (CONFIG.defaultMeetingMinutes || 30) * 60000);
  }

  function capitalize(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

  function formatRange(start, end, allDay) {
    var optD = { weekday: 'short', month: 'short', day: 'numeric' };
    var optT = { hour: 'numeric', minute: '2-digit' };
    if (allDay) return start.toLocaleDateString([], optD) + ' (all day)';
    var s = start.toLocaleDateString([], optD) + ' ' + start.toLocaleTimeString([], optT);
    if (!end) return s;
    var sameDay = start.toDateString() === end.toDateString();
    var e = sameDay ? end.toLocaleTimeString([], optT)
                    : end.toLocaleDateString([], optD) + ' ' + end.toLocaleTimeString([], optT);
    return s + ' \u2013 ' + e;
  }

  function openModal(m) { show(m); }
  function closeModal(m) { hide(m); }

  function showToast(text, link) {
    var t = els.toast;
    t.textContent = '';
    var span = document.createElement('span'); span.textContent = text; t.appendChild(span);
    if (link && /^https:\/\//i.test(link)) {
      var a = document.createElement('a');
      a.href = link; a.target = '_blank'; a.rel = 'noopener';
      a.className = 'toast-link'; a.textContent = 'Open Teams meeting';
      t.appendChild(a);
    }
    show(t);
    requestAnimationFrame(function () { t.classList.add('shown'); });
    clearTimeout(t._timer);
    t._timer = setTimeout(function () {
      t.classList.remove('shown');
      setTimeout(function () { hide(t); }, 300);
    }, link ? 8000 : 4000);
  }

  function setStatus(msg, isError, busy) {
    els.statusText.textContent = msg;
    els.statusBar.classList.toggle('error', !!isError);
    els.statusBar.classList.toggle('busy', !!busy);
    show(els.statusBar);
  }
  function clearStatus() { hide(els.statusBar); }

  function friendlyError(e) { return (e && e.message) ? e.message : String(e); }
})();
