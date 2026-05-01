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

// ── Custom Confirm Dialog ─────────────────────────────────────────────
function showConfirm(message) {
    return new Promise((resolve) => {
        document.getElementById('confirmMessage').textContent = message;
        document.getElementById('confirmOverlay').style.display = 'block';

        document.getElementById('confirmYes').onclick = () => {
            document.getElementById('confirmOverlay').style.display = 'none';
            resolve(true);
        };

        document.getElementById('confirmNo').onclick = () => {
            document.getElementById('confirmOverlay').style.display = 'none';
            resolve(false);
        };
    });
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

    document.getElementById('btnCreate').disabled = true;
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
        document.getElementById('btnCreate').disabled = false;
    }
}

// ── Load Active Tasks ─────────────────────────────────────────────────
async function loadActiveTasks(tab) {
    const selectId = tab === 'pause' ? 'cmbActiveTaskPause' : 'cmbActiveTaskClose';
    const select = document.getElementById(selectId);

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

        showStatus(tab, snapshot.size + ' active task(s) loaded.', 'success');

    } catch (error) {
        console.error('Error loading tasks:', error);
        showStatus(tab, 'Error: ' + error.message, 'error');
    }
}

// ── Load Paused Tasks ─────────────────────────────────────────────────
async function loadPausedTasks() {
    const select = document.getElementById('cmbPausedTaskResume');

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

        showStatus('resume', snapshot.size + ' paused task(s) loaded.', 'success');

    } catch (error) {
        console.error('Error loading tasks:', error);
        showStatus('resume', 'Error: ' + error.message, 'error');
    }
}

// ── Task Selected in Pause Dropdown ──────────────────────────────────
async function taskSelectedPause() {
    const taskId = document.getElementById('cmbActiveTaskPause').value;
    if (!taskId) {
        document.getElementById('txtActiveSincePause').value = '';
        return;
    }
    try {
        const doc = await db.collection('tasks').doc(taskId).get();
        if (doc.exists) {
            document.getElementById('txtActiveSincePause').value = doc.data().activateTime;
        }
    } catch (error) {
        console.error('Error fetching task:', error);
    }
}

// ── Task Selected in Resume Dropdown ─────────────────────────────────
async function taskSelectedResume() {
    const taskId = document.getElementById('cmbPausedTaskResume').value;
    if (!taskId) {
        document.getElementById('txtPausedSinceResume').value = '';
        return;
    }
    try {
        const doc = await db.collection('tasks').doc(taskId).get();
        if (doc.exists) {
            document.getElementById('txtPausedSinceResume').value = doc.data().inactivateTime;
        }
    } catch (error) {
        console.error('Error fetching task:', error);
    }
}

// ── Task Selected in Close Dropdown ──────────────────────────────────
async function taskSelectedClose() {
    const taskId = document.getElementById('cmbActiveTaskClose').value;
    if (!taskId) {
        document.getElementById('txtActiveSinceClose').value = '';
        return;
    }
    try {
        const doc = await db.collection('tasks').doc(taskId).get();
        if (doc.exists) {
            document.getElementById('txtActiveSinceClose').value = doc.data().activateTime;
        }
    } catch (error) {
        console.error('Error fetching task:', error);
    }
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
        clearPause();

    } catch (error) {
        console.error('Error pausing task:', error);
        showStatus('pause', 'Error: ' + error.message, 'error');
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
        clearResume();

    } catch (error) {
        console.error('Error resuming task:', error);
        showStatus('resume', 'Error: ' + error.message, 'error');
    }
}

// ── Close Task ────────────────────────────────────────────────────────
async function closeTask() {
    const taskId = document.getElementById('cmbActiveTaskClose').value;
    const closeTime = document.getElementById('txtDateTimeClose').value;
    const comments = document.getElementById('txtCommentsClose').value;

    if (!taskId || !closeTime) {
        showStatus('close', 'Please select a task and set close time.', 'error');
        return;
    }

    // Custom confirm dialog
    const confirmed = await showConfirm('Are you sure you want to close task ' + taskId + '? This cannot be undone.');
    if (!confirmed) return;

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
        clearClose();

    } catch (error) {
        console.error('Error closing task:', error);
        showStatus('close', 'Error: ' + error.message, 'error');
    }
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