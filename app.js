// ── Office.js Initialization ──────────────────────────────────────────
Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Task Recorder add-in loaded successfully.");
    }
});

// ── Tab Switching ─────────────────────────────────────────────────────
function showTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelectorAll('.tab').forEach(btn => {
        btn.classList.remove('active');
    });
    document.getElementById('tab-' + tabName).classList.add('active');
    event.target.classList.add('active');

    // Auto load tasks when switching tabs
    if (tabName === 'pause') {
        loadActiveTasks('pause');
    } else if (tabName === 'resume') {
        loadPausedTasks();
    } else if (tabName === 'close') {
        loadActiveTasks('close');
    }
}

// ── Extract Full Email Data (Create tab) ──────────────────────────────
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

// ── Extract Date & Time Only (Pause/Resume/Close tabs) ────────────────
function extractDateTime(tab) {
    const item = Office.context.mailbox.item;

    if (!item) {
        showStatus(tab, 'No email selected.', 'error');
        return;
    }

    let dateTime;
    if (tab === 'resume') {
        dateTime = new Date(item.dateTimeCreated);
    } else {
        dateTime = new Date(item.dateSent || item.dateTimeCreated);
    }

    const formatted = formatDateTime(dateTime);

    if (tab === 'pause') {
        document.getElementById('txtDateTimePause').value = formatted;
    } else if (tab === 'resume') {
        document.getElementById('txtDateTimeResume').value = formatted;
    } else if (tab === 'close') {
        document.getElementById('txtDateTimeClose').value = formatted;
    }
}

// ── Generate Sequential Task ID ───────────────────────────────────────
async function generateTaskId() {
    const snapshot = await db.collection('tasks')
        .orderBy('taskId', 'desc')
        .limit(1)
        .get();

    if (snapshot.empty) {
        return '0001';
    }

    const lastTask = snapshot.docs[0].data();
    const lastNumber = parseInt(lastTask.taskId, 10);
    return String(lastNumber + 1).padStart(4, '0');
}

// ── Create Task ───────────────────────────────────────────────────────
async function createTask() {
    const subject = document.getElementById('txtSubject').value;
    const sender = document.getElementById('txtSender').value;

    if (!subject || !sender) {
        showStatus('create', 'Subject and Sender are required!', 'error');
        return;
    }

    const createBtn = document.getElementById('btnCreate');
    createBtn.disabled = true;
    createBtn.textContent = 'Creating...';
    showStatus('create', 'Creating task...', 'success');

    try {
        const taskId = await generateTaskId();
        const actionDateTime = document.getElementById('txtDateTime').value;
        const comments = document.getElementById('txtComments').value;
        const taskType = document.querySelector('input[name="taskType"]:checked').value;

        const newTask = {
            taskId:          taskId,
            subject:         subject,
            senderName:      sender,
            receiverName:    document.getElementById('txtReceiver').value,
            projectId:       document.getElementById('txtProjectID').value,
            createdDateTime: actionDateTime,
            status:          'Active',
            activateTime:    actionDateTime,
            inactivateTime:  '',
            workTimeMin:     0,
            comments:        '[' + actionDateTime + '] Created - ' + comments,
            site:            document.getElementById('cmbSite').value,
            taskType:        taskType,
            priorityScore:   parseInt(document.getElementById('txtPriority').value, 10)
        };

        await db.collection('tasks').doc(taskId).set(newTask);

        showStatus('create', 'Task ' + taskId + ' created successfully! ✅', 'success');
        clearCreate();

    } catch (error) {
        console.error('Error creating task:', error);
        showStatus('create', 'Error: ' + error.message, 'error');
        createBtn.disabled = false;
        createBtn.textContent = 'Create';
    }
}

// ── Load Active Tasks ─────────────────────────────────────────────────
async function loadActiveTasks(tab) {
    const selectId = tab === 'pause' ? 'cmbActiveTaskPause' : 'cmbActiveTaskClose';
    const select = document.getElementById(selectId);

    // Disable action button while loading
    const actionBtnId = tab === 'pause' ? 'btnPause' : 'btnClose';
    const actionBtn = document.getElementById(actionBtnId);
    actionBtn.disabled = true;

    select.innerHTML = '<option value="">-- Loading... --</option>';

    try {
        const snapshot = await db.collection('tasks')
            .where('status', '==', 'Active')
            .orderBy('taskId', 'asc')
            .get();

        select.innerHTML = '<option value="">-- Select Task --</option>';

        if (snapshot.empty) {
            showStatus(tab, 'No active tasks found.', 'error');
            return;
        }

        snapshot.forEach(doc => {
            const task = doc.data();
            const option = document.createElement('option');
            option.value = task.taskId;
            option.textContent = task.taskId + ' - ' + task.subject.substring(0, 50);
            select.appendChild(option);
        });

        showStatus(tab, snapshot.size + ' active task(s) loaded. ✅', 'success');

    } catch (error) {
        console.error('Error loading tasks:', error);
        showStatus(tab, 'Error: ' + error.message, 'error');
    }
}

// ── Load Paused Tasks ─────────────────────────────────────────────────
async function loadPausedTasks() {
    const select = document.getElementById('cmbPausedTaskResume');
    const resumeBtn = document.getElementById('btnResume');
    resumeBtn.disabled = true;

    select.innerHTML = '<option value="">-- Loading... --</option>';

    try {
        const snapshot = await db.collection('tasks')
            .where('status', '==', 'Paused')
            .orderBy('taskId', 'asc')
            .get();

        select.innerHTML = '<option value="">-- Select Task --</option>';

        if (snapshot.empty) {
            showStatus('resume', 'No paused tasks found.', 'error');
            return;
        }

        snapshot.forEach(doc => {
            const task = doc.data();
            const option = document.createElement('option');
            option.value = task.taskId;
            option.textContent = task.taskId + ' - ' + task.subject.substring(0, 50);
            select.appendChild(option);
        });

        showStatus('resume', snapshot.size + ' paused task(s) loaded. ✅', 'success');

    } catch (error) {
        console.error('Error loading tasks:', error);
        showStatus('resume', 'Error: ' + error.message, 'error');
    }
}

// ── Task Selected in Pause Dropdown ──────────────────────────────────
async function taskSelectedPause() {
    const taskId = document.getElementById('cmbActiveTaskPause').value;
    const pauseBtn = document.getElementById('btnPause');

    if (!taskId) {
        document.getElementById('txtActiveSincePause').value = '';
        pauseBtn.disabled = true;
        return;
    }

    try {
        const doc = await db.collection('tasks').doc(taskId).get();
        if (doc.exists) {
            document.getElementById('txtActiveSincePause').value = doc.data().activateTime;
            pauseBtn.disabled = false;
        }
    } catch (error) {
        console.error('Error fetching task:', error);
        pauseBtn.disabled = true;
    }
}

// ── Task Selected in Resume Dropdown ─────────────────────────────────
async function taskSelectedResume() {
    const taskId = document.getElementById('cmbPausedTaskResume').value;
    const resumeBtn = document.getElementById('btnResume');

    if (!taskId) {
        document.getElementById('txtPausedSinceResume').value = '';
        resumeBtn.disabled = true;
        return;
    }

    try {
        const doc = await db.collection('tasks').doc(taskId).get();
        if (doc.exists) {
            document.getElementById('txtPausedSinceResume').value = doc.data().inactivateTime;
            resumeBtn.disabled = false;
        }
    } catch (error) {
        console.error('Error fetching task:', error);
        resumeBtn.disabled = true;
    }
}

// ── Task Selected in Close Dropdown ──────────────────────────────────
async function taskSelectedClose() {
    const taskId = document.getElementById('cmbActiveTaskClose').value;
    const closeBtn = document.getElementById('btnClose');

    if (!taskId) {
        document.getElementById('txtActiveSinceClose').value = '';
        closeBtn.disabled = true;
        return;
    }

    try {
        const doc = await db.collection('tasks').doc(taskId).get();
        if (doc.exists) {
            document.getElementById('txtActiveSinceClose').value = doc.data().activateTime;
            closeBtn.disabled = false;
        }
    } catch (error) {
        console.error('Error fetching task:', error);
        closeBtn.disabled = true;
    }
}

// ── Helper: Validate Action DateTime ─────────────────────────────────
function validateDateTime(dateTimeStr, referenceStr, tab, checkFuture = true) {
    const pattern = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/;
    if (!pattern.test(dateTimeStr)) {
        showStatus(tab, '❌ Invalid date format. Use yyyy-mm-dd hh:mm', 'error');
        return false;
    }

    const actionTime = new Date(dateTimeStr.replace(' ', 'T'));

    if (checkFuture && actionTime > new Date()) {
        showStatus(tab, '❌ Time cannot be in the future.', 'error');
        return false;
    }

    if (referenceStr) {
        const referenceTime = new Date(referenceStr.replace(' ', 'T'));
        if (actionTime <= referenceTime) {
            return false;
        }
    }

    return true;
}

// ── Pause Task ────────────────────────────────────────────────────────
async function pauseTask() {
    const taskId = document.getElementById('cmbActiveTaskPause').value;
    const pauseTime = document.getElementById('txtDateTimePause').value;
    const comments = document.getElementById('txtCommentsPause').value;

    if (!taskId || !pauseTime) {
        showStatus('pause', 'Please select a task and set pause time.', 'error');
        return;
    }

    const activeSince = document.getElementById('txtActiveSincePause').value;
    if (!validateDateTime(pauseTime, activeSince, 'pause')) {
        showStatus('pause', '❌ Pause time must be after active since: ' + activeSince, 'error');
        return;
    }

    const pauseBtn = document.getElementById('btnPause');
    pauseBtn.disabled = true;
    pauseBtn.textContent = 'Pausing...';

    try {
        const docRef = db.collection('tasks').doc(taskId);
        const doc = await docRef.get();
        const task = doc.data();

        const activateTime = new Date(task.activateTime.replace(' ', 'T'));
        const pauseDateTime = new Date(pauseTime.replace(' ', 'T'));
        const diffMinutes = Math.round((pauseDateTime - activateTime) / 60000);
        const newWorkTime = (task.workTimeMin || 0) + diffMinutes;

        const newComment = task.comments +
            '\n[' + pauseTime + '] Paused - ' + comments;

        await docRef.update({
            status:         'Paused',
            inactivateTime: pauseTime,
            workTimeMin:    newWorkTime,
            comments:       newComment
        });

        showStatus('pause', 'Task ' + taskId + ' paused. Work time: ' + newWorkTime + ' mins ✅', 'success');
        document.getElementById('cmbActiveTaskPause').innerHTML = '<option value="">-- Select Task --</option>';
        document.getElementById('txtActiveSincePause').value = '';
        clearPause();

    } catch (error) {
        console.error('Error pausing task:', error);
        showStatus('pause', 'Error: ' + error.message, 'error');
    } finally {
        pauseBtn.disabled = true;
        pauseBtn.textContent = 'Pause';
    }
}

// ── Resume Task ───────────────────────────────────────────────────────
async function resumeTask() {
    const taskId = document.getElementById('cmbPausedTaskResume').value;
    const resumeTime = document.getElementById('txtDateTimeResume').value;
    const comments = document.getElementById('txtCommentsResume').value;

    if (!taskId || !resumeTime) {
        showStatus('resume', 'Please select a task and set resume time.', 'error');
        return;
    }

    const pausedSince = document.getElementById('txtPausedSinceResume').value;
    if (!validateDateTime(resumeTime, pausedSince, 'resume')) {
        showStatus('resume', '❌ Resume time must be after paused since: ' + pausedSince, 'error');
        return;
    }

    const resumeBtn = document.getElementById('btnResume');
    resumeBtn.disabled = true;
    resumeBtn.textContent = 'Resuming...';

    try {
        const docRef = db.collection('tasks').doc(taskId);
        const doc = await docRef.get();
        const task = doc.data();

        const newComment = task.comments +
            '\n[' + resumeTime + '] Resumed - ' + comments;

        await docRef.update({
            status:         'Active',
            activateTime:   resumeTime,
            inactivateTime: '',
            comments:       newComment
        });

        showStatus('resume', 'Task ' + taskId + ' resumed. Clock restarted ✅', 'success');
        document.getElementById('cmbPausedTaskResume').innerHTML = '<option value="">-- Select Task --</option>';
        document.getElementById('txtPausedSinceResume').value = '';
        clearResume();

    } catch (error) {
        console.error('Error resuming task:', error);
        showStatus('resume', 'Error: ' + error.message, 'error');
    } finally {
        resumeBtn.disabled = true;
        resumeBtn.textContent = 'Resume';
    }
}

// ── Close Task ────────────────────────────────────────────────────────
let closeConfirmPending = false;
let closeConfirmTimer = null;

async function closeTask() {
    const taskId = document.getElementById('cmbActiveTaskClose').value;
    const closeTime = document.getElementById('txtDateTimeClose').value;
    const comments = document.getElementById('txtCommentsClose').value;

    if (!taskId || !closeTime) {
        showStatus('close', 'Please select a task and set close time.', 'error');
        return;
    }

    // First click — ask for confirmation
    if (!closeConfirmPending) {
        closeConfirmPending = true;

        const btn = document.getElementById('btnClose');
        btn.style.backgroundColor = '#c0392b';
        btn.textContent = 'Close Task';

        showStatus('close', '⚠️ Click Close again to confirm. Auto cancels in 10 seconds.', 'error');

        closeConfirmTimer = setTimeout(() => {
            resetCloseButton();
        }, 10000);

        return;
    }

    // Second click — validate then proceed
    clearTimeout(closeConfirmTimer);
    resetCloseButton();

    const activeSinceClose = document.getElementById('txtActiveSinceClose').value;
    if (!validateDateTime(closeTime, activeSinceClose, 'close')) {
        showStatus('close', '❌ Close time must be after active since: ' + activeSinceClose, 'error');
        return;
    }

    const closeBtn = document.getElementById('btnClose');
    closeBtn.disabled = true;
    closeBtn.textContent = 'Closing...';

    try {
        const docRef = db.collection('tasks').doc(taskId);
        const doc = await docRef.get();
        const task = doc.data();

        const activateTime = new Date(task.activateTime.replace(' ', 'T'));
        const closeDateTime = new Date(closeTime.replace(' ', 'T'));
        const diffMinutes = Math.round((closeDateTime - activateTime) / 60000);
        const finalWorkTime = (task.workTimeMin || 0) + diffMinutes;
        const finalWorkHours = (finalWorkTime / 60).toFixed(1);

        const newComment = task.comments +
            '\n[' + closeTime + '] Closed - ' + comments;

        await docRef.update({
            status:         'Closed',
            inactivateTime: closeTime,
            workTimeMin:    finalWorkTime,
            comments:       newComment
        });

        showStatus('close', 'Task ' + taskId + ' closed. Total: ' + finalWorkTime + ' mins (' + finalWorkHours + ' hrs) ✅', 'success');
        document.getElementById('cmbActiveTaskClose').innerHTML = '<option value="">-- Select Task --</option>';
        document.getElementById('txtActiveSinceClose').value = '';
        clearClose();

    } catch (error) {
        console.error('Error closing task:', error);
        showStatus('close', 'Error: ' + error.message, 'error');
    } finally {
        closeBtn.disabled = true;
        closeBtn.textContent = 'Close';
    }
}

function resetCloseButton() {
    closeConfirmPending = false;
    const btn = document.getElementById('btnClose');
    btn.style.backgroundColor = '';
    btn.textContent = 'Close';
    showStatus('close', '', '');
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
    document.getElementById('txtPriority').value = '5';
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
        setTimeout(() => {
            statusEl.className = 'status-msg';
        }, 4000);
    }
}
