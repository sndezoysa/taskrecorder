// ── Office.js Initialization ──────────────────────────────────────────
Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Task Recorder add-in loaded successfully.");
    }
});

// ── Tab Switching ─────────────────────────────────────────────────────
function showTab(tabName) {
    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });

    // Remove active from all tab buttons
    document.querySelectorAll('.tab').forEach(btn => {
        btn.classList.remove('active');
    });

    // Show selected tab content
    document.getElementById('tab-' + tabName).classList.add('active');

    // Mark selected tab button as active
    event.target.classList.add('active');
}

// ── Extract Email Data ────────────────────────────────────────────────
function extractEmail() {
    const item = Office.context.mailbox.item;

    if (!item) {
        showStatus('create', 'No email selected.', 'error');
        return;
    }

    document.getElementById('txtSubject').value = item.subject || '';
    document.getElementById('txtSender').value = item.from.displayName || '';
    document.getElementById('txtReceiver').value = item.to.map(r => r.displayName).join(', ') || '';

    const received = new Date(item.dateTimeCreated);
    document.getElementById('txtDateTime').value = formatDateTime(received);

    document.getElementById('btnCreate').disabled = false;
}

// ── Create Task (placeholder - Firebase comes in Phase 3) ─────────────
function createTask() {
    const subject = document.getElementById('txtSubject').value;
    const sender = document.getElementById('txtSender').value;

    if (!subject || !sender) {
        showStatus('create', 'Subject and Sender are required!', 'error');
        return;
    }

    // Placeholder - will connect to Firebase in Phase 3
    showStatus('create', 'Firebase not connected yet. Coming in Phase 3!', 'error');
}

// ── Pause Task (placeholder) ──────────────────────────────────────────
function pauseTask() {
    showStatus('pause', 'Firebase not connected yet. Coming in Phase 3!', 'error');
}

// ── Resume Task (placeholder) ─────────────────────────────────────────
function resumeTask() {
    showStatus('resume', 'Firebase not connected yet. Coming in Phase 3!', 'error');
}

// ── Close Task (placeholder) ──────────────────────────────────────────
function closeTask() {
    showStatus('close', 'Firebase not connected yet. Coming in Phase 3!', 'error');
}

// ── Load Active Tasks (placeholder) ──────────────────────────────────
function loadActiveTasks(tab) {
    showStatus(tab, 'Firebase not connected yet. Coming in Phase 3!', 'error');
}

// ── Load Paused Tasks (placeholder) ──────────────────────────────────
function loadPausedTasks() {
    showStatus('resume', 'Firebase not connected yet. Coming in Phase 3!', 'error');
}

// ── Set Current Time ──────────────────────────────────────────────────
function setCurrentTime(tab) {
    const now = formatDateTime(new Date());

    if (tab === 'pause') {
        document.getElementById('txtDateTimePause').value = now;
    } else if (tab === 'resume') {
        document.getElementById('txtDateTimeResume').value = now;
    } else if (tab === 'close') {
        document.getElementById('txtDateTimeClose').value = now;
    }
}

// ── Clear Forms ───────────────────────────────────────────────────────
function clearCreate() {
    document.getElementById('txtSubject').value = '';
    document.getElementById('txtSender').value = '';
    document.getElementById('txtReceiver').value = '';
    document.getElementById('txtDateTime').value = '';
    document.getElementById('txtProjectID').value = '';
    document.getElementById('txtComments').value = '';
    document.getElementById('cmbSite').value = 'MEDK';
    document.querySelector('input[name="taskType"][value="Requested"]').checked = true;
    document.getElementById('btnCreate').disabled = true;
}

function clearPause() {
    document.getElementById('txtDateTimePause').value = '';
    document.getElementById('txtCommentsPause').value = '';
}

function clearResume() {
    document.getElementById('txtDateTimeResume').value = '';
    document.getElementById('txtCommentsResume').value = '';
}

function clearClose() {
    document.getElementById('txtDateTimeClose').value = '';
    document.getElementById('txtCommentsClose').value = '';
}

// ── Helper: Format Date ───────────────────────────────────────────────
function formatDateTime(date) {
    const pad = n => String(n).padStart(2, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

// ── Helper: Show Status Message ───────────────────────────────────────
function showStatus(tab, message, type) {
    const statusId = 'status-' + tab;
    let statusEl = document.getElementById(statusId);

    if (statusEl) {
        statusEl.textContent = message;
        statusEl.className = 'status-msg ' + type;

        // Auto hide after 4 seconds
        setTimeout(() => {
            statusEl.className = 'status-msg';
        }, 4000);
    }
}