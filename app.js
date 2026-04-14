const officerList = document.getElementById('officerList');
const loanList = document.getElementById('loanList');
const addOfficerBtn = document.getElementById('addOfficerBtn');
const addLoanBtn = document.getElementById('addLoanBtn');
const chooseFolderBtn = document.getElementById('chooseFolderBtn');
const changeFolderBtn = document.getElementById('changeFolderBtn');
const randomizeBtn = document.getElementById('randomizeBtn');
const sampleBtn = document.getElementById('sampleBtn');
const clearBtn = document.getElementById('clearBtn');
const messageEl = document.getElementById('message');
const outputStepEl = document.getElementById('outputStep');
const outputStepCompactEl = document.getElementById('outputStepCompact');
const outputStepDetailsEl = document.getElementById('outputStepDetails');
const folderStatusEl = document.getElementById('folderStatus');
const folderPromptEl = document.getElementById('folderPrompt');
const loanAssignmentsEl = document.getElementById('loanAssignments');
const officerAssignmentsEl = document.getElementById('officerAssignments');

let outputDirectoryHandle = null;

function createInputRow(type, value = '') {
  const row = document.createElement('div');
  row.className = 'row';

  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = type === 'officer' ? 'Loan officer name' : 'Loan name or ID';
  input.value = value;

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.textContent = '×';
  removeBtn.className = 'remove-btn';
  removeBtn.addEventListener('click', () => row.remove());

  row.appendChild(input);
  row.appendChild(removeBtn);
  return row;
}

function addOfficer(value = '') {
  officerList.appendChild(createInputRow('officer', value));
}

function addLoan(value = '') {
  loanList.appendChild(createInputRow('loan', value));
}

function getValues(container) {
  return [...container.querySelectorAll('input')]
    .map((input) => input.value.trim())
    .filter(Boolean);
}

function shuffle(array) {
  const copy = [...array];
  for (let i = copy.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function chooseRandom(items, count) {
  return shuffle(items).slice(0, count);
}

function setMessage(text = '', tone = 'warning') {
  messageEl.textContent = text;
  messageEl.dataset.tone = text ? tone : '';
}

function supportsFolderSelection() {
  return typeof window.showDirectoryPicker === 'function';
}

function updateFolderStatus() {
  if (outputDirectoryHandle) {
    folderStatusEl.textContent = `Selected folder: ${outputDirectoryHandle.name}`;
    folderStatusEl.dataset.state = 'ready';
    outputStepEl.dataset.state = 'complete';
    outputStepCompactEl.hidden = false;
    outputStepDetailsEl.hidden = true;
    randomizeBtn.disabled = false;
    randomizeBtn.dataset.state = 'ready';
    return;
  }

  if (!supportsFolderSelection()) {
    folderPromptEl.textContent = 'Folder selection is not supported in this browser. Use a current version of Microsoft Edge or Google Chrome.';
    folderPromptEl.dataset.state = 'error';
    outputStepEl.dataset.state = 'error';
    outputStepCompactEl.hidden = true;
    outputStepDetailsEl.hidden = false;
    randomizeBtn.disabled = true;
    randomizeBtn.dataset.state = 'locked';
    return;
  }

  folderPromptEl.textContent = 'No output folder selected.';
  folderPromptEl.dataset.state = 'idle';
  outputStepEl.dataset.state = 'pending';
  outputStepCompactEl.hidden = true;
  outputStepDetailsEl.hidden = false;
  randomizeBtn.disabled = true;
  randomizeBtn.dataset.state = 'locked';
}

async function ensureDirectoryPermission(directoryHandle) {
  if (typeof directoryHandle.queryPermission !== 'function') {
    return true;
  }

  const options = { mode: 'readwrite' };
  if (await directoryHandle.queryPermission(options) === 'granted') {
    return true;
  }

  return (await directoryHandle.requestPermission(options)) === 'granted';
}

async function chooseOutputFolder() {
  if (!supportsFolderSelection()) {
    setMessage('Choose Output Folder is only available in browsers that support folder access, such as current Microsoft Edge or Google Chrome.', 'warning');
    updateFolderStatus();
    return;
  }

  try {
    const directoryHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    const hasPermission = await ensureDirectoryPermission(directoryHandle);

    if (!hasPermission) {
      setMessage('Folder access was not granted. Please choose a folder and allow write access.', 'warning');
      return;
    }

    outputDirectoryHandle = directoryHandle;
    updateFolderStatus();
    setMessage(`Output folder selected: ${directoryHandle.name}`, 'success');
  } catch (error) {
    if (error.name === 'AbortError') {
      return;
    }

    setMessage(`Unable to select an output folder: ${error.message}`, 'warning');
  }
}

function padNumber(value) {
  return String(value).padStart(2, '0');
}

function buildPdfFileName(date) {
  const year = date.getFullYear();
  const month = padNumber(date.getMonth() + 1);
  const day = padNumber(date.getDate());
  const hours = padNumber(date.getHours());
  const minutes = padNumber(date.getMinutes());
  const seconds = padNumber(date.getSeconds());

  return `Loan-Randomized-Results-${year}-${month}-${day}-${hours}${minutes}${seconds}.pdf`;
}

function formatDisplayTimestamp(date) {
  return new Intl.DateTimeFormat(undefined, {
    year: 'numeric',
    month: 'long',
    day: '2-digit',
    hour: 'numeric',
    minute: '2-digit',
    second: '2-digit'
  }).format(date);
}

function buildPdfLines(result, officers, loans, generatedAt) {
  const lines = [
    { text: 'Loan Randomized Results', size: 18, gapAfter: 18 },
    { text: `Generated: ${formatDisplayTimestamp(generatedAt)}`, size: 11, gapAfter: 6 },
    { text: `Loan officers entered: ${officers.length}`, size: 11, gapAfter: 4 },
    { text: `Loans entered: ${loans.length}`, size: 11, gapAfter: 14 },
    { text: 'Assignments by Loan', size: 14, gapAfter: 10 }
  ];

  result.loanAssignments.forEach((entry) => {
    lines.push({ text: `${entry.loan} -> ${entry.officers[0]}`, size: 11, gapAfter: 6 });
  });

  lines.push({ text: '', size: 11, gapAfter: 8 });
  lines.push({ text: 'Assignments by Officer', size: 14, gapAfter: 10 });

  Object.entries(result.officerAssignments).forEach(([officer, assignedLoans]) => {
    lines.push({ text: `${officer} (${assignedLoans.length})`, size: 12, gapAfter: 6 });

    if (!assignedLoans.length) {
      lines.push({ text: 'No loans assigned.', size: 11, indent: 16, gapAfter: 6 });
      return;
    }

    assignedLoans.forEach((loan) => {
      lines.push({ text: `- ${loan}`, size: 11, indent: 16, gapAfter: 5 });
    });

    lines.push({ text: '', size: 11, gapAfter: 4 });
  });

  return lines;
}

function writePdfLines(doc, lines) {
  const pageHeight = doc.internal.pageSize.getHeight();
  const maxWidth = 500;
  const left = 54;
  const top = 64;
  const bottom = 54;
  let currentY = top;

  lines.forEach((line) => {
    const fontSize = line.size || 11;
    const indent = line.indent || 0;
    const text = line.text || ' ';
    const wrappedLines = doc.splitTextToSize(text, maxWidth - indent);
    const lineHeight = fontSize + 4;
    const requiredHeight = wrappedLines.length * lineHeight + (line.gapAfter || 0);

    if (currentY + requiredHeight > pageHeight - bottom) {
      doc.addPage();
      currentY = top;
    }

    doc.setFont('helvetica', fontSize >= 14 ? 'bold' : 'normal');
    doc.setFontSize(fontSize);
    doc.text(wrappedLines, left + indent, currentY);
    currentY += wrappedLines.length * lineHeight + (line.gapAfter || 0);
  });
}

async function saveResultPdf(result, officers, loans, generatedAt) {
  if (!outputDirectoryHandle) {
    throw new Error('No output folder has been selected.');
  }

  if (!window.jspdf || !window.jspdf.jsPDF) {
    throw new Error('The PDF library did not load correctly.');
  }

  const doc = new window.jspdf.jsPDF({ unit: 'pt', format: 'letter' });
  writePdfLines(doc, buildPdfLines(result, officers, loans, generatedAt));

  const pdfBlob = doc.output('blob');
  const fileName = buildPdfFileName(generatedAt);
  const fileHandle = await outputDirectoryHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();

  await writable.write(pdfBlob);
  await writable.close();

  return fileName;
}

function assignLoans(officers, loans) {
  const cleanOfficers = [...new Set(officers.map((name) => name.trim()).filter(Boolean))];
  const cleanLoans = loans.map((loan) => loan.trim()).filter(Boolean);

  const loanCount = cleanLoans.length;
  const officerCount = cleanOfficers.length;

  if (officerCount < 1) {
    return { error: 'Please add at least one loan officer.' };
  }

  if (officerCount !== officers.length) {
    return { error: 'Loan officer names must be unique so assignments are tracked correctly.' };
  }

  if (!loanCount) {
    return { error: 'Please add at least one loan.' };
  }

  const officerAssignments = {};
  cleanOfficers.forEach((officer) => {
    officerAssignments[officer] = [];
  });

  const loanAssignments = [];
  const shuffledLoans = shuffle(cleanLoans);

  // Build randomized assignment slots in fair rounds.
  // Example with 2 officers and 5 loans:
  // round 1 => [Ashley, Jade] in random order
  // round 2 => [Ashley, Jade] in new random order
  // round 3 => [Ashley] or [Jade] depending on the random order
  const assignmentSlots = [];
  while (assignmentSlots.length < loanCount) {
    assignmentSlots.push(...shuffle(cleanOfficers));
  }

  const selectedSlots = assignmentSlots.slice(0, loanCount);

  shuffledLoans.forEach((loan, index) => {
    const assignedOfficer = selectedSlots[index];
    officerAssignments[assignedOfficer].push(loan);
    loanAssignments.push({
      loan,
      officers: [assignedOfficer],
      shared: false
    });
  });

  return { loanAssignments, officerAssignments };
}

function renderResults(result) {
  if (result.error) {
    setMessage(result.error, 'warning');
    loanAssignmentsEl.className = 'results empty';
    officerAssignmentsEl.className = 'results empty';
    loanAssignmentsEl.textContent = 'No assignments yet.';
    officerAssignmentsEl.textContent = 'No assignments yet.';
    return;
  }

  setMessage('');

  loanAssignmentsEl.className = 'results';
  officerAssignmentsEl.className = 'results';

  loanAssignmentsEl.innerHTML = '';
  officerAssignmentsEl.innerHTML = '';

  result.loanAssignments.forEach((entry) => {
    const div = document.createElement('div');
    div.className = 'loan-line';

    if (entry.shared) {
      div.innerHTML = `
        <div><span class="assignment-name">${escapeHtml(entry.loan)}</span></div>
        <div class="shared">Shared across: ${entry.officers.map(escapeHtml).join(', ')}</div>
      `;
    } else {
      div.innerHTML = `
        <div><span class="assignment-name">${escapeHtml(entry.loan)}</span></div>
        <div>Assigned to: ${escapeHtml(entry.officers[0])}</div>
      `;
    }

    loanAssignmentsEl.appendChild(div);
  });

  Object.entries(result.officerAssignments).forEach(([officer, assignedLoans]) => {
    const group = document.createElement('div');
    group.className = 'result-group';

    const badge = `<span class="badge">${assignedLoans.length} assigned</span>`;
    group.innerHTML = `<h3>${escapeHtml(officer)} ${badge}</h3>`;

    if (!assignedLoans.length) {
      const empty = document.createElement('div');
      empty.className = 'hint';
      empty.textContent = 'No loans assigned.';
      group.appendChild(empty);
    } else {
      assignedLoans.forEach((loan) => {
        const pill = document.createElement('span');
        pill.className = 'loan-pill';
        pill.textContent = loan;
        group.appendChild(pill);
      });
    }

    officerAssignmentsEl.appendChild(group);
  });
}

function escapeHtml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

addOfficerBtn.addEventListener('click', () => addOfficer());
addLoanBtn.addEventListener('click', () => addLoan());
chooseFolderBtn.addEventListener('click', () => {
  chooseOutputFolder();
});
changeFolderBtn.addEventListener('click', () => {
  chooseOutputFolder();
});

randomizeBtn.addEventListener('click', async () => {
  if (!outputDirectoryHandle) {
    setMessage('Choose an output folder before randomizing assignments.', 'warning');
    updateFolderStatus();
    return;
  }

  const officers = getValues(officerList);
  const loans = getValues(loanList);
  const result = assignLoans(officers, loans);
  renderResults(result);

  if (result.error) {
    return;
  }

  try {
    const generatedAt = new Date();
    const fileName = await saveResultPdf(result, officers, loans, generatedAt);
    setMessage(`Assignments randomized and saved to ${fileName}.`, 'success');
  } catch (error) {
    setMessage(`Assignments were generated, but the PDF could not be saved: ${error.message}`, 'warning');
  }
});

sampleBtn.addEventListener('click', () => {
  officerList.innerHTML = '';
  loanList.innerHTML = '';

  ['Alex', 'Brooke', 'Chris', 'Dana'].forEach(addOfficer);
  ['Loan 101', 'Loan 102', 'Loan 103', 'Loan 104', 'Loan 105'].forEach(addLoan);

  const result = assignLoans(getValues(officerList), getValues(loanList));
  renderResults(result);
});

clearBtn.addEventListener('click', () => {
  officerList.innerHTML = '';
  loanList.innerHTML = '';
  setMessage('');
  loanAssignmentsEl.className = 'results empty';
  officerAssignmentsEl.className = 'results empty';
  loanAssignmentsEl.textContent = 'No assignments yet.';
  officerAssignmentsEl.textContent = 'No assignments yet.';
  addOfficer();
  addOfficer();
  addOfficer();
  addOfficer();
});

addOfficer('Loan Officer 1');
addOfficer('Loan Officer 2');
addOfficer('Loan Officer 3');
addOfficer('Loan Officer 4');
addLoan('Loan A');
addLoan('Loan B');
updateFolderStatus();
