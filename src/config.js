// Outlook Scheduler configuration.
// Fill in clientId and tenantId after running deploy/register-app.ps1 (it prints both).
window.OSA_CONFIG = {
  clientId: '6e91bb81-2ef3-4af0-8a27-a85d392dbed3',
  tenantId: '769d2378-c654-4f20-ad01-3f0192c35475',

  // Leave null to auto-detect "<this folder>/blank.html" as the MSAL redirect URI.
  // Set explicitly only if you host blank.html somewhere else.
  redirectUri: null,

  // Delegated Microsoft Graph scopes (read own calendar + create events with Teams meetings).
  scopes: ['User.Read', 'Calendars.ReadWrite'],

  graphBase: 'https://graph.microsoft.com/v1.0',

  // Default length (minutes) for a new appointment when the agent clicks "New appointment".
  defaultMeetingMinutes: 30,

  // Also create a Dynamics 365 appointment activity (subject, time, Teams link) when booking.
  // Requires the app to be embedded in a model-driven app (uses Xrm.WebApi).
  createDynamicsAppointment: true,

  // Manager view: allow looking up a colleague's availability (free/busy via Graph getSchedule).
  allowViewOthers: true,

  // Roster the manager can pick calendars from. Replace these with real user UPNs for live
  // availability — the sample names below are fictitious and drive the demo/preview.
  teamMembers: [
    { name: 'Sara Nguyen', email: 'sara.nguyen@contoso.com' },
    { name: 'Diego Ramirez', email: 'diego.ramirez@contoso.com' },
    { name: 'Priya Patel', email: 'priya.patel@contoso.com' },
    { name: 'Tom Becker', email: 'tom.becker@contoso.com' },
    { name: 'Amina Hassan', email: 'amina.hassan@contoso.com' },
    { name: 'Liam O\u2019Brien', email: 'liam.obrien@contoso.com' }
  ],

  // Working-hours window shown on the calendar grid.
  businessHours: { start: '08:00', end: '18:00' }
};

var CONFIG = window.OSA_CONFIG;
