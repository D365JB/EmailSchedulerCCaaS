// Optional: also create a Dynamics 365 appointment activity when booking, so the meeting is
// tracked in Dataverse (timeline/activities) with its scheduled time and Teams join link.
// Uses Xrm.WebApi, so it only works when the app is embedded in a model-driven app.
// getXrm() is defined in auth.js (loaded first).

function dataverseAvailable() {
  var xrm = getXrm();
  return !!(xrm && xrm.WebApi && xrm.WebApi.createRecord);
}

// opts: { subject, start(Date), end(Date), notes, location, joinUrl, attendeeEmail, regardingEntity, regardingId }
async function createDynamicsAppointment(opts) {
  var xrm = getXrm();
  if (!xrm || !xrm.WebApi || !xrm.WebApi.createRecord) {
    throw new Error('Not running inside a model-driven app.');
  }

  var parts = [];
  if (opts.notes) parts.push(opts.notes);
  if (opts.attendeeEmail) parts.push('Attendee: ' + opts.attendeeEmail);
  if (opts.joinUrl) parts.push('Join Microsoft Teams meeting: ' + opts.joinUrl);

  var record = {
    subject: opts.subject,
    scheduledstart: opts.start.toISOString(),
    scheduledend: opts.end.toISOString(),
    location: opts.joinUrl ? 'Microsoft Teams Meeting' : (opts.location || ''),
    description: parts.join('\n\n')
  };

  // Link to a record (e.g. the active case/contact) when passed via ?regardingtype= & ?regardingid=.
  if (opts.regardingEntity && opts.regardingId) {
    record['regardingobjectid_' + opts.regardingEntity + '@odata.bind'] =
      '/' + entitySetName(opts.regardingEntity) + '(' + opts.regardingId + ')';
  }

  return xrm.WebApi.createRecord('appointment', record);
}

function entitySetName(logicalName) {
  var map = { incident: 'incidents', contact: 'contacts', account: 'accounts', lead: 'leads', opportunity: 'opportunities' };
  return map[logicalName] || (logicalName + 's');
}
