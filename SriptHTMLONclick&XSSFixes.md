You are fixing security and runtime bugs in Script.html of a Google Apps Script HR web app.

The project uses this helper: const jsq = v => JSON.stringify(v).replace(/"/g, '&quot;');
This helper safely encodes any value for use inside HTML attributes.

Fix the following bugs:

─── BUG A: Toast XSS (line 222) ───
CURRENT:
  toast.innerHTML = `<span>${icons[type] || icons.default}</span><span>${message}</span>`;
FIX:
  Replace innerHTML with DOM construction:
  toast.textContent = '';
  const iconSpan = document.createElement('span');
  iconSpan.textContent = icons[type] || icons.default;
  const msgSpan = document.createElement('span');
  msgSpan.textContent = message;
  toast.appendChild(iconSpan);
  toast.appendChild(msgSpan);

─── BUG B: JSON.stringify entire candidate object in onclick (around line 1419) ───
CURRENT:
  onclick="Router.navigate('candidate-detail',{selectedCandidate:${JSON.stringify(c).replace(/"/g, '&quot;')}})"
FIX:
  onclick="Router.navigate('candidate-detail',{selectedCandidateId:${jsq(c.CandidateID)}})"

  Then in the candidateDetail view renderer, update the line that reads App.state.selectedCandidate:
  CURRENT: const c = App.state.selectedCandidate;
  FIX:
    const c = App.state.selectedCandidateId
      ? (App.state.candidates || []).find(x => x.CandidateID === App.state.selectedCandidateId)
      : App.state.selectedCandidate;
    if (!c) { Router.navigate('candidates'); return; }

─── BUG C: All onclick/onsubmit handlers using old single-quote interpolation ───
Replace ALL of the following occurrences (use jsq() for every value):

Line 1533: onclick="Views._openEditProfileModal('${escHtml(c.CandidateID)}')"
→ onclick="Views._openEditProfileModal(${jsq(c.CandidateID)})"

Line 1536: onclick="Views._openStatusModal('${escHtml(c.CandidateID)}','${escHtml(c.FullName)}')"
→ onclick="Views._openStatusModal(${jsq(c.CandidateID)},${jsq(c.FullName)})"

Line 1539: onclick="Views._openEventModal('${escHtml(c.CandidateID)}')"
→ onclick="Views._openEventModal(${jsq(c.CandidateID)})"

Line 1542: onclick="Views._openUploadModal('${escHtml(c.CandidateID)}','${escHtml(c.DriveFolderID || '')}')"
→ onclick="Views._openUploadModal(${jsq(c.CandidateID)},${jsq(c.DriveFolderID || '')})"

Line 1762: onclick="Views._submitStatus('${candidateId}')"
→ onclick="Views._submitStatus(${jsq(candidateId)})"

Line 1809: onclick="Views._submitEditProfile('${candidateId}')"
→ onclick="Views._submitEditProfile(${jsq(candidateId)})"

Line 1885: onclick="Views._submitUpload('${candidateId}','${folderId}')"
→ onclick="Views._submitUpload(${jsq(candidateId)},${jsq(folderId)})"

Line 2448: onclick="Views._markEventComplete('${e.EventID}')"
→ onclick="Views._markEventComplete(${jsq(e.EventID)})"

Line 2474: onsubmit="Views._submitEvent(event, '${candidateId}')"
→ onsubmit="Views._submitEvent(event, ${jsq(candidateId)})"

After making all changes, commit with message:
"fix: onclick XSS & single-quote interpolation vulnerabilities in Script.html"