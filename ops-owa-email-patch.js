// Ops Comms – Email compose (OWA-first) patch
// - Opens Outlook on the web compose with full body where possible
// - If body is very long, opens with concise body and copies FULL (HTML + text)
//   to clipboard so you can paste formatting
// - Fixes email validation regex and recipient normalisation

// --- Utilities ---
function isWhitbreadEmail(email){
  const e = String(email || '').trim();
  if(!e) return false;
  // firstname.lastname@whitbread.com or @whitbread.co.uk
  const re = /^[A-Za-z]+\.[A-Za-z0-9\-]+@whitbread\.(com|co\.uk)$/i;
  return re.test(e);
}
function anyActionSelected(){
  return ['act_threat','act_remove','act_licence','act_allergy','act_pos','act_project','act_other']
    .some(id => document.getElementById(id)?.checked);
}
function autoCriticality(){
  const threat = document.getElementById('act_threat').checked;
  const remove = document.getElementById('act_remove').checked;
  const licence = document.getElementById('act_licence').checked;
  const allergy = document.getElementById('act_allergy').checked;
  const pos = document.getElementById('act_pos').checked;
  const project = document.getElementById('act_project').checked;
  const other = document.getElementById('act_other').checked;
  if(threat) return 'critical';
  if(remove || licence || allergy) return 'important';
  if(pos || project || other) return 'other';
  return 'other';
}
function selectedCriticality(){
  if(document.getElementById('crit_critical').checked) return 'critical';
  if(document.getElementById('crit_important').checked) return 'important';
  if(document.getElementById('crit_other').checked) return 'other';
  return 'auto';
}
function criticalityLabel(c){
  if(c==='critical') return 'Critical (4 hours)';
  if(c==='important') return 'Important (same day at 23:59)';
  return 'Other (>24 hours)';
}
function hoursFor(c){
  if(c==='critical') return 4;
  if(c==='important') return 24; // unused for EOD, kept for completeness
  return 72; // default
}
function updateCritPill(c){
  const pill = document.getElementById('crit_pill');
  if(!pill) return;
  pill.classList.remove('critical','important','other');
  pill.classList.add(c==='critical'?'critical':(c==='important'?'important':'other'));
  pill.textContent = `Criticality: ${criticalityLabel(c)}`;
}
function toLocalInputValue(date){
  const pad = n => String(n).padStart(2,'0');
  const y = date.getFullYear();
  const m = pad(date.getMonth()+1);
  const d = pad(date.getDate());
  const h = pad(date.getHours());
  const min = pad(date.getMinutes());
  return `${y}-${m}-${d}T${h}:${min}`;
}
function fromLocalInputValue(val){
  if(!val) return null;
  const [dpart, tpart] = val.split('T');
  const [y,m,d] = dpart.split('-').map(Number);
  const [hh,mm] = tpart.split(':').map(Number);
  return new Date(y, m-1, d, hh, mm);
}
function endOfDay(date){ return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 0, 0); }

// --- Email helpers (OWA deeplink + clipboard & .eml fallback) ---
function escapeHtml(text){
  return String(text || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function textToHtml(text){
  return escapeHtml(text).replace(/(?:\r\n|\r|\n)/g,'<br>');
}
function normalizeRecipients(listStr){
  return String(listStr||'')
    .replace(/[\n\r]/g, ',')
    .replace(/;/g, ',')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .join(', ');
}
function buildOwaHref({to='',cc='',bcc='',subject='',bodyText=''}) {
  const parts = [];
  if (to) parts.push('to=' + encodeURIComponent(normalizeRecipients(to)));
  // cc/bcc intentionally omitted to keep URL short
  if (subject) parts.push('subject=' + encodeURIComponent(subject));
  if (bodyText) parts.push('body=' + encodeURIComponent(bodyText));
  return 'https://outlook.office.com/mail/deeplink/compose?' + parts.join('&');
}
async function copyFullToClipboard({ plain, html }){
  try {
    if (navigator.clipboard && window.ClipboardItem) {
      const items = { 'text/plain': new Blob([plain], {type:'text/plain'}) };
      if (html) items['text/html'] = new Blob([html], {type:'text/html'});
      await navigator.clipboard.write([ new ClipboardItem(items) ]);
      return true;
    }
  } catch(e) { /* fall through */ }
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(plain);
      return true;
    }
  } catch(e) { /* ignore */ }
  return false;
}
function tryOpenOwaThenFallback({ fullBody, shortBody, subject, to, fullHtml }) {
  const MAX_SAFE = 11000; // generous for modern browsers
  let href = buildOwaHref({ to, subject, bodyText: fullBody });
  if (href.length <= MAX_SAFE) {
    window.open(href, '_blank', 'noopener');
    // copy full (HTML+text) in background
    copyFullToClipboard({ plain: fullBody, html: fullHtml }).then(ok => {
      if (ok) {
        alert('Email opened in Outlook on the web.\n\nThe full, formatted summary is on your clipboard. Paste into the email to keep bold section headers.');
      } else {
        alert('Email opened in Outlook on the web.\n\nIf you need the full summary, copy it from the “Task summary” block and paste into the email.');
      }
    });
    return { mode:'owa-full' };
  }
  // Too long for full, try short body
  href = buildOwaHref({ to, subject, bodyText: shortBody });
  if (href.length <= MAX_SAFE) {
    copyFullToClipboard({ plain: fullBody, html: fullHtml });
    window.open(href, '_blank', 'noopener');
    return { mode:'owa-short' };
  }
  return { mode:'too-long' };
}
function downloadEmlDraft({ to='', cc='', bcc='', subject='', bodyHtml='' }){
  const allTo = [to, cc, bcc].filter(Boolean).join(', ');
  const headers = [];
  if (allTo) headers.push(`To: ${allTo}`);
  headers.push(`Subject: ${subject}`);
  headers.push('MIME-Version: 1.0');
  headers.push('Content-Type: text/html; charset=UTF-8');
  const eml = headers.join('\r\n') + '\r\n\r\n' + bodyHtml;
  const blob = new Blob([eml], { type: 'message/rfc822' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const safe = (subject || 'draft').replace(/[^\w\- \[\]\(\)]+/g,' ').trim() || 'draft';
  a.download = `${safe}.eml`;
  a.click();
  URL.revokeObjectURL(a.href);
}

let currentEmail = null;

// --- Core generator ---
function computeDeadline(){
  const schedVal = document.getElementById('schedule_dt').value;
  const override = document.getElementById('deadline_override').checked;
  const critSel = selectedCriticality();
  const crit = (critSel==='auto') ? autoCriticality() : critSel;
  updateCritPill(crit);
  const sched = fromLocalInputValue(schedVal);
  if(!override && sched){
    let deadline;
    if(crit==='critical'){
      deadline = new Date(sched.getTime() + 4*60*60*1000);
    } else if(crit==='important'){
      deadline = endOfDay(sched); // same day 23:59
    } else {
      deadline = new Date(sched.getTime() + hoursFor(crit)*60*60*1000);
    }
    document.getElementById('deadline_dt').value = toLocalInputValue(deadline);
  }
  const weekendEl = document.getElementById('weekend');
  if(sched){
    const day = sched.getDay(); // 0=Sun, 6=Sat
    weekendEl.style.display = (day===0 || day===6) ? 'block' : 'none';
  } else {
    weekendEl.style.display = 'none';
  }
}
function onCritChange(){ computeDeadline(); }
function onScheduleChange(){ computeDeadline(); }
function whoSelected(){
  if(document.getElementById('who_admin').checked) return 'Site Admin (Site Managers)';
  if(document.getElementById('who_checker').checked) return 'Compliance Checker (All other site roles)';
  return null;
}

function generate(){
  computeDeadline();
  const title = document.getElementById('title').value.trim();
  const reqName = document.getElementById('req_name').value.trim();
  const reqEmail = document.getElementById('req_email').value.trim();
  const desc = document.getElementById('desc').value.trim();
  const purpose = document.getElementById('purpose').value.trim();
  const outcomeText = document.getElementById('outcome_text').value.trim();
  const audience = document.getElementById('audience').value.trim();
  const dependencies = document.getElementById('dependencies').value.trim();
  const essential = document.getElementById('chk_essential').checked;
  const reinforce = document.getElementById('chk_reinforce').checked;
  const escalation = document.getElementById('escalation').value.trim();
  const policy = document.getElementById('policy').value;
  const schedVal = document.getElementById('schedule_dt').value;
  const deadlineVal = document.getElementById('deadline_dt').value;
  const who = whoSelected();
  const siteListAttach = document.getElementById('site_list_attach').checked;

  let critSel = selectedCriticality();
  let crit = (critSel==='auto') ? autoCriticality() : critSel;
  updateCritPill(crit);
  const sched = fromLocalInputValue(schedVal);
  const deadline = fromLocalInputValue(deadlineVal);

  // decision logic
  let reasons = [];
  let decision = 'review';
  if(!essential){
    decision = 'declined';
    reasons.push('Core principle missing: not essential to record/track safety, compliance, or service quality.');
  }
  if(!reinforce){
    decision = 'declined';
    reasons.push('Core principle missing: task does not reinforce actions & expectations / duplicates comms details.');
  }
  if(!isWhitbreadEmail(escalation)){
    if(decision!=='declined') decision = 'review';
    reasons.push('Escalation contact must be a named Whitbread email.');
  }
  if(policy==='no'){
    decision = 'declined';
    reasons.push('Action not aligned to policy/standard procedure. Validate before sending.');
  }
  if(policy==='unsure'){
    if(decision!=='declined') decision = 'review';
    reasons.push('Policy alignment is unsure – validation required before sending.');
  }
  if(crit==='other'){
    if(decision!=='declined') decision = 'review';
    reasons.push('"Other" criticality (>24h) requires approval for task applicability and lead time.');
  }
  if(sched && deadline){
    const diffHrs = Math.round((deadline.getTime()-sched.getTime())/3600000);
    if(crit==='critical'){
      if(diffHrs>4){ reasons.push('Deadline is longer than recommended for selected criticality.'); }
      if(diffHrs<4){
        if(decision!=='declined') decision = 'review';
        reasons.push('Deadline shorter than recommended – ensure urgency is justified.');
      }
    } else if(crit==='important'){
      const eod = endOfDay(sched);
      if(deadline.getTime() > eod.getTime()){
        reasons.push('Deadline is later than the end-of-day default for Important.');
      }
    }
  }
  if(decision!=='declined' && essential && reinforce && isWhitbreadEmail(escalation) && policy==='yes' && (crit==='critical' || crit==='important')){
    decision = 'approved';
  }

  // Missing list
  const missing = [];
  if(!title) missing.push('Title');
  if(!reqName) missing.push('Requester name');
  if(!reqEmail){
    missing.push('Requester email');
  } else {
    const emailValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(reqEmail);
    if(!emailValid) missing.push('Requester email (invalid format)');
  }
  if(!isWhitbreadEmail(escalation)) missing.push('Escalation contact (named Whitbread email)');
  if(!anyActionSelected()) missing.push('Select at least one action type');
  if(!essential) missing.push('Confirm essential to record/track safety/compliance/service quality');
  if(!reinforce) missing.push('Confirm task reinforces actions & expectations (no duplication)');
  if(!sched) missing.push('Schedule date/time');
  if(!deadline) missing.push('Deadline date/time');
  if(policy==='unsure') missing.push('Validate policy alignment (currently Unsure)');
  if(!who) missing.push('Select who can complete the task (Site Admin / Compliance Checker)');
  if(!siteListAttach) missing.push('Confirm a site list will be attached to the request email');
  if(!desc) missing.push('Task wording (including actions)');
  if(!purpose) missing.push('Purpose of the request');
  if(!outcomeText) missing.push('Desired outcome / what success looks like');
  if(!audience) missing.push('Impacted audience / scope');
  if(!dependencies) missing.push('Dependencies or prerequisites (if any)');

  // Render Missing and Decision UI
  const missingEl = document.getElementById('missing');
  if (missingEl) missingEl.innerHTML = missing.length ? missing.map(m=>`<li>${m}</li>`).join('') : '<li>None – all core details present.</li>';
  const missingWrap = document.getElementById('missing_wrap');
  const missingTitle = document.getElementById('missing_title');
  if (missingTitle) missingTitle.textContent = `Missing information / to confirm${missing.length ? ' ('+missing.length+')' : ''}`;
  if (missingWrap) { missingWrap.classList.remove('flash'); void missingWrap.offsetWidth; missingWrap.classList.add('flash'); }
  const decisionBox = document.getElementById('decision');
  if (decisionBox){
    decisionBox.classList.remove('approved','review','declined');
    decisionBox.classList.add(decision);
    decisionBox.querySelector('h3').innerHTML = `<span>Decision:</span> ${decision==='approved'?'Approved to go in Co Pilot':(decision==='declined'?'Declined – revise and validate':'Review required – seek approval/validation')}`;
    const reasonsEl = document.getElementById('reasons');
    if (reasonsEl) reasonsEl.innerHTML = reasons.map(r=>`<li>${r}</li>`).join('');
  }

  const fmt =(d)=> d? d.toLocaleString('en-GB', {dateStyle:'medium', timeStyle:'short'}) : '[not set]';
  const delta = (sched && deadline) ? ` (Δ ${Math.round((deadline.getTime()-sched.getTime())/3600000)}h)` : '';

  // Action labels
  const actionLines = [];
  ['act_threat','act_remove','act_licence','act_allergy','act_pos','act_project','act_other'].forEach(id=>{
    const el = document.getElementById(id);
    if(el?.checked){
      const label = el.parentElement.querySelector('span').textContent.trim();
      actionLines.push(`- ${label}`);
    }
  });
  if(actionLines.length===0){ actionLines.push('- [Select at least one action type]'); }

  // Build subject, recipients, intro
  let subject = '';
  let to = '';
  let cc = '';
  if(decision==='approved' || decision==='review'){
    subject = `Co Pilot Ops Comms Task Request – ${title || '[Title]'} (${decision==='approved'?'Approved':'Approval Required'})`;
    to = (decision==='approved' ? 'copilot.project@whitbread.com' : 'gillian.klarin@whitbread.com');
    cc = 'mark.fewtrell@whitbread.com, luke.wheeler@whitbread.com';
  }

  // Build PLAIN summary lines (for on-page preview and for body text)
  const lines = [];
  if(decision==='approved' || decision==='review'){
    const allTo = [to, cc].filter(Boolean).join(', ');
    const intro = (decision==='approved')
      ? 'Below is a summary of an approved request for an Ops Comms'
      : 'I would like to request approval for the below task summary to be considered as an Ops Comms.';
    lines.push(`Subject: ${subject}`);
    lines.push(`To: ${allTo}`);
    lines.push('');
    lines.push('Hello,');
    lines.push('');
    lines.push(intro);
    lines.push('');
    lines.push('Summary of request');
    lines.push('Title: ' + (title || '[Enter title]'));
    lines.push(`Requester: ${reqName || '[Name]'} (${reqEmail || '[Email]'})`);
    lines.push(`Escalation contact: ${escalation || '[Named Whitbread email]'}`);
    lines.push(`Who can complete: ${who || '[Select role]'}`);
    lines.push(`Site list attached: ${siteListAttach ? 'Yes' : 'No'}`);
    lines.push('');
    lines.push('Purpose & objectives');
    lines.push(`Purpose: ${purpose || '[Explain why this is needed]'}`);
    lines.push(`Outcome (what success looks like): ${outcomeText || '[Define the expected end state]'}`);
    lines.push(`Impacted audience / scope: ${audience || '[Who is affected]'}`);
    lines.push(`Dependencies / prerequisites: ${dependencies || '[List any dependencies or approvals]'}`);
    lines.push('');
    lines.push('Action type(s):');
    lines.push(...actionLines);
    lines.push('');
    lines.push(`Criticality: ${criticalityLabel(crit)}`);
    lines.push(`Schedule: ${fmt(sched)}  Deadline: ${fmt(deadline)}${delta}`);
    lines.push('');
    lines.push('Decision & rationale');
    lines.push(`${decision==='approved'?'Approved to go in Co Pilot': 'Approval required – seek validation'}`);
    if(reasons.length){
      lines.push('Reasons:');
      reasons.forEach(r=>lines.push(`- ${r}`));
    } else {
      lines.push('Reasons: Meets core principles and policy.');
    }
    lines.push('');
    lines.push('Missing information / to confirm');
    const missingListForSummary = missing.length ? missing : ['None – all core details present.'];
    missingListForSummary.forEach(m=>lines.push(`- ${m}`));
    lines.push('');
    lines.push('Task wording (including actions)');
    lines.push(desc || '[Provide the task wording and actions]');
  } else {
    lines.push('Decision: Declined – revise and validate');
    if(reasons.length){
      lines.push('Reasons:');
      reasons.forEach(r=>lines.push(`- ${r}`));
    }
    lines.push('');
    lines.push('Missing information / to confirm');
    const missingListForSummary = missing.length ? missing : ['None – all core details present.'];
    missingListForSummary.forEach(m=>lines.push(`- ${m}`));
    lines.push('');
    lines.push('Purpose & objectives');
    lines.push(`Purpose: ${purpose || '[Explain why this is needed]'}`);
    lines.push(`Outcome (what success looks like): ${outcomeText || '[Define the expected end state]'}`);
    lines.push(`Impacted audience / scope: ${audience || '[Who is affected]'}`);
    lines.push(`Dependencies / prerequisites: ${dependencies || '[List any dependencies or approvals]'}`);
    lines.push('');
    lines.push('Task wording (including actions)');
    lines.push(desc || '[Provide the task wording and actions]');
  }

  // Show the summary on the page
  const summaryEl = document.getElementById('summary');
  if (summaryEl) summaryEl.textContent = lines.join('\n');

  // Build HTML version with bold section headers (for clipboard / .eml fallback)
  function section(titleHtml, innerHtml){
    return `<p><strong>${titleHtml}</strong></p>\n${innerHtml}`;
  }
  function list(items){ return `<ul>` + items.map(x=>`<li>${escapeHtml(x)}</li>`).join('') + `</ul>`; }
  const htmlParts = [];
  if(decision==='approved' || decision==='review'){
    const intro = (decision==='approved')
      ? 'Below is a summary of an approved request for an Ops Comms'
      : 'I would like to request approval for the below task summary to be considered as an Ops Comms.';
    htmlParts.push('<p>Hello,</p>');
    htmlParts.push(`<p>${escapeHtml(intro)}.</p>`);
    htmlParts.push(section('Summary of request',
      list([
        `Title: ${title || '[Enter title]'}`,
        `Requester: ${reqName || '[Name]'} (${reqEmail || '[Email]'})`,
        `Escalation contact: ${escalation || '[Named Whitbread email]'}`,
        `Who can complete: ${who || '[Select role]'}`,
        `Site list attached: ${siteListAttach ? 'Yes' : 'No'}`
      ])
    ));
    htmlParts.push(section('Purpose &amp; objectives',
      list([
        `Purpose: ${purpose || '[Explain why this is needed]'}`,
        `Outcome (what success looks like): ${outcomeText || '[Define the expected end state]'}`,
        `Impacted audience / scope: ${audience || '[Who is affected]'}`,
        `Dependencies / prerequisites: ${dependencies || '[List any dependencies or approvals]'}`
      ])
    ));
    htmlParts.push(section('Action type(s)', list(actionLines)));
    htmlParts.push(section('Timing & criticality', list([
      `Criticality: ${criticalityLabel(crit)}`,
      `Schedule: ${fmt(sched)}  Deadline: ${fmt(deadline)}${delta}`
    ])));
    const decInner = [];
    decInner.push(escapeHtml(decision==='approved'?'Approved to go in Co Pilot': 'Approval required – seek validation'));
    if(reasons.length){ decInner.push('<p><strong>Reasons:</strong></p>' + list(reasons)); }
    htmlParts.push(section('Decision & rationale', decInner.join('\n')));
    const missItems = missing.length ? missing : ['None – all core details present.'];
    htmlParts.push(section('Missing information / to confirm', list(missItems)));
    htmlParts.push(section('Task wording (including actions)', `<div style="white-space:pre-wrap">${escapeHtml(desc || '[Provide the task wording and actions]')}</div>`));
  } else {
    htmlParts.push(section('Decision', 'Declined – revise and validate'));
    if(reasons.length){ htmlParts.push('<p><strong>Reasons:</strong></p>' + list(reasons)); }
    const missItems = missing.length ? missing : ['None – all core details present.'];
    htmlParts.push(section('Missing information / to confirm', list(missItems)));
    htmlParts.push(section('Purpose &amp; objectives', list([
      `Purpose: ${purpose || '[Explain why this is needed]'}`,
      `Outcome (what success looks like): ${outcomeText || '[Define the expected end state]'}`,
      `Impacted audience / scope: ${audience || '[Who is affected]'}`,
      `Dependencies / prerequisites: ${dependencies || '[List any dependencies or approvals]'}`
    ])));
    htmlParts.push(section('Task wording (including actions)', `<div style="white-space:pre-wrap">${escapeHtml(desc || '[Provide the task wording and actions]')}</div>`));
  }
  const htmlBody = htmlParts.join('\n');

  // Prepare email compose data
  const btn = document.getElementById('gen_email');
  if (decision === 'declined') {
    currentEmail = null;
    if (btn) { btn.disabled = true; btn.title = 'Decision is Declined – no email generated'; }
  } else {
    const allTo = [to, cc].filter(Boolean).join(', ');
    const fullBody = lines.slice(2).join('\r\n'); // CRLF
    const idxMissing = lines.findIndex(l => l.trim().toLowerCase().startsWith('missing information'));
    const shortSliceEnd = idxMissing > -1 ? idxMissing : Math.min(lines.length, 40);
    const shortBodyCore = lines.slice(2, shortSliceEnd);
    const shortBody = [
      ...shortBodyCore,
      '',
      '—',
      'Note: Full summary has been copied to your clipboard. Paste it here if you need the full details.'
    ].join('\r\n');

    currentEmail = { to: allTo, cc: '', bcc: '', subject, fullBody, shortBody, htmlBody };
    if (btn) { btn.disabled = false; btn.title = 'Generate email draft'; }
  }

  // UX
  const outcomeCard = document.getElementById('outcome');
  if (outcomeCard) outcomeCard.scrollIntoView({ behavior:'smooth', block:'start' });
  const inner = document.getElementById('outcome_inner');
  if (inner){ inner.classList.remove('flash'); void inner.offsetWidth; inner.classList.add('flash'); }
}

function copySummary(){
  const txt = document.getElementById('summary').textContent;
  navigator.clipboard.writeText(txt).then(()=>{ alert('Summary copied to clipboard'); })
  .catch(()=>{ alert('Copy failed. Select and copy manually.'); });
}
function downloadTxt(){
  const blob = new Blob([document.getElementById('summary').textContent], {type: 'text/plain'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'cp_task_decision_summary.txt';
  a.click();
}
function resetForm(){
  document.querySelectorAll('input, textarea, select').forEach(el=>{
    if(el.type==='checkbox'){ el.checked=false; }
    else if(el.type==='radio'){ el.checked=false; }
    else if(el.tagName==='SELECT'){ el.selectedIndex=0; }
    else{ el.value=''; }
  });
  document.getElementById('crit_auto').checked = true;
  updateCritPill('other');
  document.getElementById('decision').className = 'decision review';
  document.getElementById('decision').querySelector('h3').innerHTML = '<span>Decision:</span> Review required';
  document.getElementById('reasons').innerHTML = '';
  document.getElementById('missing').innerHTML = '';
  document.getElementById('missing_title').textContent = 'Missing information / to confirm';
  document.getElementById('summary').textContent='';
  document.getElementById('weekend').style.display = 'none';
  const btn = document.getElementById('gen_email');
  if (btn) { btn.disabled = true; btn.title = 'Generate a decision first'; }
}

// Wire-up (expecting same DOM IDs as your v3 file)
document.getElementById('generate')?.addEventListener('click', generate);
document.getElementById('generate_bottom')?.addEventListener('click', generate);
document.getElementById('copy')?.addEventListener('click', copySummary);
document.getElementById('download')?.addEventListener('click', downloadTxt);
document.getElementById('reset')?.addEventListener('click', resetForm);
document.getElementById('deadline_override')?.addEventListener('change', (e)=>{
  document.getElementById('deadline_dt').readOnly = !e.target.checked;
  if(!e.target.checked){ computeDeadline(); }
});
['crit_auto','crit_critical','crit_important','crit_other'].forEach(id=>{
  document.getElementById(id)?.addEventListener('change', onCritChange);
});
document.getElementById('schedule_dt')?.addEventListener('change', onScheduleChange);

// Generate Email flow with OWA-first, clipboard, then .eml fallback
document.getElementById('gen_email')?.addEventListener('click', ()=>{
  if (!currentEmail) { alert('No email is available for this decision.'); return; }
  const { to, subject, fullBody, shortBody, htmlBody } = currentEmail;
  const result = tryOpenOwaThenFallback({ fullBody, shortBody, subject, to, fullHtml: htmlBody });
  if (result.mode === 'too-long') {
    // Final fallback: .eml download with HTML body (bold headings preserved)
    downloadEmlDraft({ to, cc:'', bcc:'', subject, bodyHtml: htmlBody || textToHtml(fullBody) });
    alert('The email content was too long for the Outlook on the web deeplink, so a draft .eml was downloaded instead.');
  }
});

// initial
updateCritPill('other');
document.getElementById('deadline_dt').readOnly = true;
const initBtn = document.getElementById('gen_email');
if (initBtn) { initBtn.disabled = true; initBtn.title = 'Generate a decision first'; }
