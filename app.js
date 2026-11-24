// Letter Management System - Cloud Version (Refactored)
const DB_NAME = 'LMS_DB';
const DB_VERSION = 1;
let db_local;

// Departments List
const departments = [
  "Administrator", "Wasiirka (Main)", "Wasiir Ku-xigeenka 1aad", "Wasiir Ku-xigeenka 2aad",
  "HRM", "Adeega Shacabka", "ICT", "Sharciyada", "Xadaynka",
  "Public Relation", "Planing", "Wacyigalinta",
  "Finance", "arkiviya-1", "arkiviya-2",
  "Tax Accounting Team", "baadhista canshurta", "Xogta canshurta", "Sirdonka canshurta", "CRM"
];

const DEFAULT_PASS = "123456";

// Hybrid Initialization
async function initDB() {
  // Local DB Init
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onerror = e => reject("DB Error: " + e.target.error);
    request.onsuccess = e => { db_local = e.target.result; resolve(db_local); };
    request.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('letters')) db.createObjectStore('letters', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('settings')) db.createObjectStore('settings');
    };
  });
}

// Data Handling - Prefers Cloud if available
// Data Handling - Real-time & Hybrid
let unsubscribeListener = null;



// Helper: Refresh whatever view user is on
function refreshCurrentView() {
  if (typeof currentView === 'undefined') return;
  if (currentView === 'dashboard') showDashboard();
  else if (currentView === 'incoming') showIncoming();
  else if (currentView === 'outgoing') showOutgoing();
  else if (currentView === 'reports') showReports();
}

// Updated getLetters - Now relies on Local DB (which is fed by Listener)
async function getLetters() {
  const allLetters = await new Promise((resolve) => {
    const transaction = db_local.transaction(['letters'], 'readonly');
    const store = transaction.objectStore('letters');
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => resolve([]);
  });

  // Filter for view
  if (currentDept === 'Administrator') return allLetters;
  return allLetters.filter(l => l.sender === currentDept || l.recipient === currentDept);
}

// MANUAL SYNC: Force a download from cloud
async function forceDataSync() {
  const btn = document.getElementById('syncBtn');
  if (btn) { btn.innerText = "Syncing..."; btn.disabled = true; }

  try {
    if (typeof db_cloud !== 'undefined' && db_cloud) {
      console.log("🔄 Force Sync Started...");
      const snapshot = await db_cloud.collection('letters').get(); // Direct fetch
      const cloudLetters = snapshot.docs.map(doc => doc.data());

      if (cloudLetters.length > 0) {
        const transaction = db_local.transaction(['letters'], 'readwrite');
        const store = transaction.objectStore('letters');
        cloudLetters.forEach(l => store.put(l));

        transaction.oncomplete = () => {
          alert(`✅ Synced! Found ${cloudLetters.length} letters from Cloud.`);
          refreshCurrentView();
          if (btn) { btn.innerText = "Sync Now"; btn.disabled = false; }
        };
      } else {
        alert("☁️ Cloud is reachable but empty.");
        if (btn) { btn.innerText = "Sync Now"; btn.disabled = false; }
      }
    } else {
      alert("❌ Cloud Database Object Missing!");
    }
  } catch (e) {
    alert("❌ Sync Failed: " + e.message);
    if (btn) { btn.innerText = "Retry"; btn.disabled = false; }
  }
}

// Helper for timeout
const timeout = (ms) => new Promise((_, reject) => setTimeout(() => reject(new Error("Request Timed Out")), ms));

async function saveLetter(letter) {
  // CLOUD FIRST STRATEGY
  const statusEl = document.getElementById('statusMessage'); // Optional UI hook

  try {
    if (typeof db_cloud !== 'undefined' && db_cloud) {
      await syncToCloud(letter); // Must succeed to continue
    } else {
      throw new Error("No Cloud Connection");
    }

    // Local Save (Optimistic Update)
    const transaction = db_local.transaction(['letters'], 'readwrite');
    const store = transaction.objectStore('letters');
    store.put(letter); // Save the object (with fileBlob for local cache if needed)

  } catch (e) {
    console.error("Save Failed:", e);
    alert("❌ MESSAGE FAILED! \n\nCould not send to cloud. Please check internet connection.\nError: " + e.message);
    throw e;
  }
}

// Helper: Convert File to Base64
const fileToBase64 = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.readAsDataURL(file);
  reader.onload = () => resolve(reader.result);
  reader.onerror = error => reject(error);
});

async function syncToCloud(letterOriginal) {
  // FINAL "NO CARD" STRATEGY: CHUNKED FIRESTORE STORAGE
  const letterPayload = {
    id: letterOriginal.id,
    sender: letterOriginal.sender,
    recipient: letterOriginal.recipient,
    subject: letterOriginal.subject,
    date: letterOriginal.date,
    status: letterOriginal.status,
    fileName: letterOriginal.fileName || null,
    fileType: letterOriginal.fileBlob ? letterOriginal.fileBlob.type : null,
    fileChunks: 0,
    timestamp: letterOriginal.timestamp || Date.now()
  };

  // 1. DATA UPLOAD (Blocking)
  if (letterOriginal.fileBlob && typeof db_cloud !== 'undefined' && db_cloud) {
    try {
      const submitBtn = document.querySelector('#newLetterModal .btn-primary');
      if (submitBtn) submitBtn.innerHTML = `<i class='material-icons-round animate-spin'>data_usage</i> Chunking Data...`;

      const base64String = await fileToBase64(letterOriginal.fileBlob);
      const chunkSize = 400 * 1024; // 400KB chunks
      const totalChunks = Math.ceil(base64String.length / chunkSize);
      letterPayload.fileChunks = totalChunks;
      letterOriginal.fileChunks = totalChunks; // PERSIST LOCALLY

      const batch = db_cloud.batch();
      for (let i = 0; i < totalChunks; i++) {
        const start = i * chunkSize;
        const end = start + chunkSize;
        const chunkData = base64String.substring(start, end);
        const chunkRef = db_cloud.collection('letters').doc(letterPayload.id).collection('chunks').doc(String(i));
        batch.set(chunkRef, { data: chunkData, index: i });
      }
      if (submitBtn) submitBtn.innerHTML = `<i class='material-icons-round animate-spin'>cloud_upload</i> Syncing Chunks...`;
      await batch.commit();
      console.log(`✅ Uploaded ${totalChunks} chunks`);
    } catch (e) {
      console.error("Data Upload Error:", e);
      throw new Error("Local File Error: " + e.message);
    }
  }

  // 2. SAVE METADATA
  try {
    await db_cloud.collection('letters').doc(letterPayload.id).set(letterPayload);
  } catch (e) { throw new Error("Registry Error: " + e.message); }
}




async function updateLetterStatus(id, newStatus, extraFields = {}) {
  // Cloud Update
  try {
    if (typeof db_cloud !== 'undefined' && db_cloud) {
      await db_cloud.collection('letters').doc(id).update({ status: newStatus, ...extraFields });
    }
  } catch (e) {
    console.warn("Cloud Sync Error (updateStatus):", e);
  }

  // Local Update (Optimistic)
  const letters = await getLetters();
  const letter = letters.find(l => l.id === id);
  if (letter) {
    letter.status = newStatus;
    Object.assign(letter, extraFields);
    const transaction = db_local.transaction(['letters'], 'readwrite');
    transaction.objectStore('letters').put(letter);
  }
  refreshCurrentView(); // Force refresh to show new status
  updateSidebarNotifications();
}

async function getPasswords() {
  try {
    if (typeof db_cloud !== 'undefined' && db_cloud) {
      // Prioritize cloud for global sync on multi-computer
      const doc = await Promise.race([
        db_cloud.collection('settings').doc('passwords').get(),
        timeout(5000)
      ]);
      if (doc.exists) {
        const cloudPasses = doc.data();
        // Sync cloud to local for later offline use
        const transaction = db_local.transaction(['settings'], 'readwrite');
        transaction.objectStore('settings').put(cloudPasses, 'passwords');
        return cloudPasses;
      }
    }
  } catch (e) {
    console.warn("Cloud Passwords Fetch Error:", e);
  }

  return new Promise((resolve) => {
    const transaction = db_local.transaction(['settings'], 'readonly');
    const store = transaction.objectStore('settings');
    const request = store.get('passwords');
    request.onsuccess = () => {
      let passes = request.result || {};
      departments.forEach(dept => { if (!passes[dept]) passes[dept] = DEFAULT_PASS; });
      resolve(passes);
    };
  });
}

async function savePasswords(passes) {
  if (typeof db_cloud !== 'undefined' && db_cloud) {
    await db_cloud.collection('settings').doc('passwords').set(passes);
  }
  const transaction = db_local.transaction(['settings'], 'readwrite');
  transaction.objectStore('settings').put(passes, 'passwords');
}

async function cleanupExpiredLetters() {
  const letters = await getLetters();
  const settings = await getSettings();
  const retentionDays = settings.retentionDays || 7;
  if (retentionDays === -1) return; // Never delete

  const RETENTION_MS = retentionDays * 24 * 60 * 60 * 1000;
  const now = Date.now();
  letters.forEach(async (l) => {
    if ((now - (l.timestamp || 0)) > RETENTION_MS) {
      await deleteLetter(l.id);
    }
  });
}

async function getSettings() {
  try {
    if (typeof db_cloud !== 'undefined' && db_cloud) {
      const doc = await db_cloud.collection('settings').doc('config').get();
      if (doc.exists) return doc.data();
    }
  } catch (e) {
    console.warn("Cloud Sync Error (getSettings):", e);
  }
  return new Promise((resolve) => {
    const transaction = db_local.transaction(['settings'], 'readonly');
    const store = transaction.objectStore('settings');
    const request = store.get('config');
    request.onsuccess = () => resolve(request.result || { retentionDays: 7 });
  });
}

async function saveSettings(settings) {
  try {
    if (typeof db_cloud !== 'undefined' && db_cloud) {
      await db_cloud.collection('settings').doc('config').set(settings);
    }
  } catch (e) {
    console.warn("Cloud Sync Error (saveSettings):", e);
  }
  const transaction = db_local.transaction(['settings'], 'readwrite');
  transaction.objectStore('settings').put(settings, 'config');
}

async function deleteLetter(id) {
  try {
    if (typeof db_cloud !== 'undefined' && db_cloud) {
      await db_cloud.collection('letters').doc(id).delete();
    }
  } catch (e) {
    console.warn("Cloud Sync Error (deleteLetter):", e);
  }
  return new Promise((resolve) => {
    const transaction = db_local.transaction(['letters'], 'readwrite');
    const store = transaction.objectStore('letters');
    store.delete(id);
    transaction.oncomplete = () => resolve();
  });
}

// State Management
let currentDept = sessionStorage.getItem('lms_current_dept') || null;

// UI Components & Rendering (Same as before but with Cloud Status)
function updateActiveNavLink(index) {
  const items = document.querySelectorAll('.nav-item');
  items.forEach((item, i) => {
    if (i === index) item.classList.add('active');
    else item.classList.remove('active');
  });
}

async function showDashboard() {
  currentView = 'dashboard';
  updateActiveNavLink(0);
  const mainBody = document.getElementById('mainBody');
  const letters = await getLetters();
  const isAdmin = currentDept === 'Administrator';

  const myIncoming = isAdmin ? letters : letters.filter(l => l.recipient === currentDept);
  const myOutgoing = isAdmin ? letters : letters.filter(l => l.sender === currentDept);
  const pendingInbox = myIncoming.filter(l => l.status === 'new' || l.status === 'pending');
  const processedCount = isAdmin ? letters.filter(l => l.status === 'processed').length : [...myIncoming, ...myOutgoing].filter(l => l.status === 'processed').length;

  const cloudStatus = (typeof db_cloud !== 'undefined' && db_cloud)
    ? '<span style="color:#22c55e; font-size:0.75rem; display:flex; align-items:center; gap:4px;"><i class="material-icons-round" style="font-size:1rem;">cloud_done</i> Cloud Active</span>'
    : '<span style="color:#f59e0b; font-size:0.75rem; display:flex; align-items:center; gap:4px;"><i class="material-icons-round" style="font-size:1rem;">cloud_off</i> Local Mode</span>';

  const mySentSummons = letters.filter(l => l.sender === currentDept && l.isSummon);
  const summonSection = mySentSummons.length > 0 ? `
    <div class="card animate-fade-in" style="border-left: 5px solid #ef4444; margin-bottom: 2.5rem; padding: 1.5rem 2rem;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:1.25rem;">
        <div>
          <h3 style="margin:0; display:flex; align-items:center; gap:8px; color:#991b1b; font-size:1.15rem; font-weight:800;">
            <i class="material-icons-round" style="color:#ef4444; font-size:1.4rem;">campaign</i>
            Live Emergency Summons Tracker
          </h3>
          <p style="margin:2px 0 0; font-size:0.8rem; color:#64748b;">Real-time attendance & acknowledgment status of departments you summoned</p>
        </div>
        <button class="btn btn-emergency" onclick="openSummonModal()" style="font-size:0.8rem; padding:0.4rem 0.9rem; background:linear-gradient(135deg,#ef4444,#dc2626); color:white; border:none; font-weight:700; border-radius:6px; cursor:pointer;">+ Summon Another</button>
      </div>
      <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 1rem;">
        ${mySentSummons.sort((a,b)=>b.timestamp-a.timestamp).slice(0, 4).map(s => `
          <div style="background: ${s.status === 'acknowledged' ? '#f0fdf4' : '#fff5f5'}; border: 1.5px solid ${s.status === 'acknowledged' ? '#bbf7d0' : '#fecaca'}; border-radius:10px; padding:1rem; box-shadow:0 2px 5px rgba(0,0,0,0.03);">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.4rem;">
              <strong style="color:var(--navy); font-size:0.95rem;">To: ${s.recipient}</strong>
              <span style="font-size:0.7rem; font-weight:800; padding:2px 8px; border-radius:4px; background:${s.status === 'acknowledged' ? '#16a34a' : '#f59e0b'}; color:white; display:inline-flex; align-items:center; gap:3px;">
                ${s.status === 'acknowledged' ? '🟢 ON THE WAY' : '🟡 WAITING'}
              </span>
            </div>
            <div style="font-size:0.85rem; color:#475569; font-weight:600;">"${s.message || 'Please come as soon as possible'}"</div>
            <div style="font-size:0.75rem; color:#64748b; margin-top:0.4rem; display:flex; justify-content:space-between;">
              <span>${s.location ? `📍 ${s.location}` : ''}</span>
              <span>${s.date}</span>
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  ` : '';

  mainBody.innerHTML = `
    <div class="animate-fade-in">
      <div style="margin-bottom: 2.5rem; display:flex; justify-content:space-between; align-items:flex-start;">
        <div>
          <h1 style="font-size: 2.25rem; font-weight: 800; color: var(--navy); margin-bottom: 0.5rem;">${isAdmin ? 'Global System Control' : 'Executive Overview'}</h1>
          <p style="color: var(--text-muted); font-size: 1.1rem;">Welcome, <strong>${currentDept}</strong>.</p>
        </div>
        <div style="background:#f8fafc; padding:8px 12px; border-radius:12px; border:1px solid #e2e8f0;">${cloudStatus}</div>
      </div>
      
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 1.5rem; margin-bottom: 3rem;">
        <div class="card" style="border-left: 5px solid #3b82f6; display: flex; align-items: center; gap: 1.5rem;">
          <div style="background: #eff6ff; color: #3b82f6; width: 60px; height: 60px; border-radius: 16px; display: flex; align-items: center; justify-content: center;"><i class="material-icons-round" style="font-size: 2rem;">send</i></div>
          <div><div style="color: #64748b; font-size: 0.85rem; font-weight: 700; text-transform: uppercase;">Outgoing</div><div style="font-size: 1.75rem; font-weight: 800;">${myOutgoing.length}</div></div>
        </div>
        <div class="card" style="border-left: 5px solid #22c55e; display: flex; align-items: center; gap: 1.5rem;">
          <div style="background: #f0fdf4; color: #22c55e; width: 60px; height: 60px; border-radius: 16px; display: flex; align-items: center; justify-content: center;"><i class="material-icons-round" style="font-size: 2rem;">mark_email_unread</i></div>
          <div><div style="color: #64748b; font-size: 0.85rem; font-weight: 700; text-transform: uppercase;">Incoming</div><div style="font-size: 1.75rem; font-weight: 800;">${myIncoming.length}</div></div>
        </div>
        <div class="card" style="border-left: 5px solid #f59e0b; display: flex; align-items: center; gap: 1.5rem;">
          <div style="background: #fffbeb; color: #f59e0b; width: 60px; height: 60px; border-radius: 16px; display: flex; align-items: center; justify-content: center;"><i class="material-icons-round" style="font-size: 2rem;">pending_actions</i></div>
          <div><div style="color: #64748b; font-size: 0.85rem; font-weight: 700; text-transform: uppercase;">Pending</div><div style="font-size: 1.75rem; font-weight: 800;">${pendingInbox.length}</div></div>
        </div>
      </div>

      ${summonSection}

      <div style="display: grid; grid-template-columns: 2fr 1fr; gap: 2rem;">
        <div class="card" style="padding: 0; overflow: hidden;">
          <div style="padding: 1.5rem 2rem; border-bottom: 1px solid #f1f5f9; display: flex; justify-content: space-between; align-items: center;"><h3 style="margin:0;">Recent Activity</h3><span style="font-size: 0.75rem; background: #f1f5f9; padding: 4px 10px; border-radius: 20px;">Last 5 Records</span></div>
          <div style="padding: 1rem 0;">
            ${[...myIncoming, ...myOutgoing].sort((a, b) => b.timestamp - a.timestamp).slice(0, 5).map(l => `
              <div style="padding: 1.25rem 2rem; display: flex; align-items: center; border-bottom: 1px solid #f8fafc;">
                <div style="width: 40px; height: 40px; border-radius: 10px; background: ${l.recipient === currentDept ? '#f0fdf4' : '#eff6ff'}; display: flex; align-items: center; justify-content: center; margin-right: 1.25rem;">
                  <i class="material-icons-round" style="color: ${l.recipient === currentDept ? '#22c55e' : '#3b82f6'};">${l.recipient === currentDept ? 'call_received' : 'call_made'}</i>
                </div>
                <div style="flex: 1;">
                  <div style="font-weight: 600;">${l.id} ${l.isSummon ? '<span style="background:linear-gradient(135deg,#ef4444,#dc2626);color:white;padding:2px 8px;border-radius:4px;font-size:0.65rem;font-weight:800;margin-left:5px;box-shadow:0 2px 4px rgba(239,68,68,0.4);">🚨 SUMMON</span>' : (l.status === 'new' && l.recipient === currentDept ? '<span style="background:#ef4444;color:white;padding:2px 6px;border-radius:4px;font-size:0.6rem;margin-left:5px;">NEW</span>' : '')}</div>
                  <div style="font-size: 0.85rem; color: #64748b;">${l.recipient === currentDept ? `From ${l.sender}` : `To ${l.recipient}`}: ${l.subject}</div>
                  ${(l.fileBlob || l.fileUrl) ? `<button class="btn btn-primary" style="padding: 2px 8px; font-size: 0.7rem; margin-top: 5px;" onclick="downloadFile('${l.id}')"><i class="material-icons-round" style="font-size: 0.9rem; vertical-align: middle;">file_download</i> Open Original</button>` : ''}
                </div>
                <div style="font-size: 0.8rem; color: #94a3b8;">${l.date}</div>
              </div>
            `).join('') || '<p style="padding: 2rem; text-align: center; color: #94a3b8;">No activity</p>'}
          </div>
        </div>
        <div class="card" style="background: linear-gradient(135deg, var(--navy) 0%, #1e293b 100%); color: white;">
          <h3 style="color: var(--accent);">Efficiency Score</h3>
          <div style="text-align: center; padding: 2rem 0;">
            <div style="font-size: 3.5rem; font-weight: 800;">${Math.round((processedCount / ([...myIncoming, ...myOutgoing].length || 1)) * 100)}%</div>
          </div>
        </div>
      </div>
    </div>
  `;
  updateSidebarNotifications();
}

async function updateSidebarNotifications() {
  const letters = await getLetters();
  if (!currentDept) return;

  // Incoming regular letters
  const incomingLetters = letters.filter(l => l.recipient === currentDept && !l.isSummon && l.status === 'new').length;
  // Incoming urgent summons
  const incomingSummons = letters.filter(l => (l.recipient === currentDept || l.recipient === 'ALL') && l.isSummon && l.status === 'urgent').length;
  
  // SENDER alerts: Pending summons sent by me waiting for arrival
  const pendingSentSummons = letters.filter(l => l.sender === currentDept && l.isSummon && l.status === 'urgent').length;

  // Sidebar badge: counts unread incoming items
  const totalIncoming = incomingLetters + incomingSummons;
  const badge = document.getElementById('incomingBadge');
  if (badge) { 
    badge.innerText = totalIncoming; 
    badge.style.display = totalIncoming > 0 ? 'flex' : 'none'; 
  }

  // Header Bell Badge: counts incoming alerts + outgoing summon activity
  const notifBadge = document.getElementById('notifBadge');
  const totalNotifs = incomingSummons + incomingLetters + pendingSentSummons;
  if (notifBadge) {
    notifBadge.innerText = totalNotifs;
    notifBadge.style.display = totalNotifs > 0 ? 'flex' : 'none';
    notifBadge.style.background = incomingSummons > 0 ? '#ef4444' : '#3b82f6';
  }
}

async function showIncoming() {
  currentView = 'incoming';
  updateActiveNavLink(1);
  const letters = await getLetters();
  for (let l of letters) { if (l.recipient === currentDept && l.status === 'new') { await updateLetterStatus(l.id, 'pending'); } }
  renderLetterTable('incoming', 'Incoming Letters');
}

function showOutgoing() {
  currentView = 'outgoing';
  updateActiveNavLink(2);
  renderLetterTable('outgoing', 'Outgoing Letters');
}

async function renderLetterTable(type, title) {
  const mainBody = document.getElementById('mainBody');
  const letters = await getLetters();
  const filtered = type === 'incoming' ? letters.filter(l => l.recipient === currentDept) : letters.filter(l => l.sender === currentDept);

  mainBody.innerHTML = `
    <div class="animate-fade-in">
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 2rem;">
        <div><h1>${title}</h1><p style="color: var(--text-muted); font-size: 0.9rem;">Total Records: ${filtered.length}</p></div>
        <div style="display:flex; gap:0.75rem; align-items:center;">
          <button class="btn btn-emergency" onclick="openSummonModal()" style="background: linear-gradient(135deg, #ef4444, #dc2626); color: white; border: none; font-weight: 700; box-shadow: 0 4px 10px rgba(239, 68, 68, 0.35); display: inline-flex; align-items: center; gap: 6px; padding: 0.55rem 1.1rem; border-radius: var(--radius-md); cursor: pointer;" title="Send Urgent Summon to a Department">
            <i class="material-icons-round" style="font-size: 1.15rem;">campaign</i>
            <span>Summon</span>
          </button>
          <button class="btn btn-primary" onclick="openNewLetterModal('${type}')"><i class="material-icons-round">add</i> Send New Letter</button>
        </div>
      </div>
      <div class="card" style="overflow-x: auto; padding: 0;">
        <table style="width: 100%; border-collapse: collapse; text-align: left;">
          <thead style="background: #f8fafc; border-bottom: 1px solid #e2e8f0;">
            <tr>
              <th style="padding: 1rem;">ID</th>
              <th style="padding: 1rem;">${type === 'incoming' ? 'Sender' : 'Recipient'}</th>
              <th style="padding: 1rem;">Subject</th>
              <th style="padding: 1rem;">Date</th>
              <th style="padding: 1rem;">Status</th>
              <th style="padding: 1rem;">Actions</th>
            </tr>
          </thead>
          <tbody>
            ${filtered.sort((a, b) => b.timestamp - a.timestamp).map(l => `
              <tr style="border-bottom: 1px solid #f1f5f9; ${l.isSummon ? 'background: #fff5f5;' : ''}">
                <td style="padding: 1rem; font-weight: 600;">
                  ${l.id}
                  ${l.isSummon ? '<span style="background:linear-gradient(135deg, #ef4444, #dc2626);color:white;padding:2px 7px;border-radius:4px;font-size:0.65rem;font-weight:800;margin-left:6px;display:inline-block;box-shadow:0 2px 4px rgba(239,68,68,0.35);">🚨 SUMMON</span>' : ''}
                </td>
                <td style="padding: 1rem;">${type === 'incoming' ? l.sender : l.recipient}</td>
                <td style="padding: 1rem; font-weight: 500;">
                  ${l.subject}
                  ${l.isSummon && l.location ? `<div style="font-size:0.75rem; color:#b91c1c; margin-top:2px;">📍 Report To: <strong>${l.location}</strong></div>` : ''}
                </td>
                <td style="padding: 1rem; color: #64748b;">${l.date}</td>
                <td style="padding: 1rem;">
                  <span style="padding: 0.3rem 0.6rem; border-radius: 99px; font-size: 0.7rem; font-weight: 700; background: ${l.isSummon && l.status === 'urgent' ? '#fee2e2' : (l.status === 'processed' || l.status === 'acknowledged' ? '#f0fdf4' : '#fffbeb')}; color: ${l.isSummon && l.status === 'urgent' ? '#b91c1c' : (l.status === 'processed' || l.status === 'acknowledged' ? '#166534' : '#92400e')};">
                    ${l.isSummon && l.status === 'urgent' ? '🚨 EMERGENCY' : (l.status === 'acknowledged' ? '✅ ACKNOWLEDGED' : l.status.toUpperCase())}
                  </span>
                </td>
                <td style="padding: 1rem;">
                  <div style="display: flex; gap: 0.5rem; align-items: center;">
                    ${(l.fileBlob || l.fileUrl) ? `
                      <button class="icon-btn" onclick="downloadFile('${l.id}')" title="Download Original File" style="color: var(--accent);">
                        <i class="material-icons-round">file_download</i>
                      </button>
                    ` : '<span style="color: #cbd5e1; font-size: 0.8rem; padding:0 0.5rem;">—</span>'}
                    <button class="icon-btn" onclick="viewLetter('${l.id}')" title="View Details">
                      <i class="material-icons-round" style="color: #64748b;">visibility</i>
                    </button>
                    <button class="icon-btn" onclick="printLetter('${l.id}')" title="Print Registry Slip">
                      <i class="material-icons-round" style="color: #64748b;">print</i>
                    </button>
                  </div>
                </td>
              </tr>
            `).join('') || '<tr><td colspan="6" style="padding:4rem; text-align:center; color:#94a3b8;"><i class="material-icons-round" style="font-size: 3rem; margin-bottom: 1rem;">inbox</i><p>No departmental records found.</p></td></tr>'}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

async function renderTableContentAsync(filtered) {
  return `
    <div class="card" style="overflow-x: auto; padding: 0;">
      <table style="width: 100%; border-collapse: collapse; text-align: left;">
        <thead style="background: #f8fafc;">
          <tr>
            <th style="padding:1rem;">ID</th>
            <th style="padding:1rem;">Subject</th>
            <th style="padding:1rem;">Actions</th>
          </tr>
        </thead>
        <tbody>
          ${filtered.map(l => `
            <tr style="border-bottom: 1px solid #f1f5f9;">
              <td style="padding:1rem; font-weight: 600;">${l.id}</td>
              <td style="padding:1rem;">${l.subject}</td>
              <td style="padding:1rem; display: flex; gap: 0.5rem;">
                ${(l.fileBlob || l.fileUrl) ? `<button class="btn btn-primary" style="padding: 0.3rem 0.6rem; font-size: 0.7rem;" onclick="downloadFile('${l.id}')">Open Document</button>` : ''}
                <button class="btn" style="padding: 0.3rem 0.6rem; font-size: 0.7rem;" onclick="viewLetter('${l.id}')">View</button>
              </td>
            </tr>
          `).join('') || '<tr><td colspan="3" style="padding:2rem; text-align:center;">No results found.</td></tr>'}
        </tbody>
      </table>
    </div>
  `;
}

async function handleSearch(query) {
  const q = query.toLowerCase();
  const letters = await getLetters();
  const filtered = letters.filter(l => (l.sender === currentDept || l.recipient === currentDept) && (l.subject.toLowerCase().includes(q) || l.id.toLowerCase().includes(q)));
  document.getElementById('mainBody').innerHTML = `<div class="animate-fade-in"><h1>Search: "${query}"</h1>${await renderTableContentAsync(filtered)}</div>`;
}

async function showReports() {
  currentView = 'reports';
  updateActiveNavLink(3);
  const mainBody = document.getElementById('mainBody');
  const letters = await getLetters();
  const myIncoming = letters.filter(l => l.recipient === currentDept);
  const myOutgoing = letters.filter(l => l.sender === currentDept);
  const processed = [...myIncoming, ...myOutgoing].filter(l => l.status === 'processed').length;

  mainBody.innerHTML = `
    <div class="animate-fade-in">
      <h1>Department Analytics</h1>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; margin-top: 2rem;">
        <div class="card"><h3>Incoming Volume</h3><p style="font-size: 2.5rem; font-weight: 800;">${myIncoming.length}</p></div>
        <div class="card"><h3>Outgoing Volume</h3><p style="font-size: 2.5rem; font-weight: 800;">${myOutgoing.length}</p></div>
      </div>
      <div class="card" style="margin-top: 2rem;"><h3>Efficiency</h3><p>Processed records: ${processed} / ${myIncoming.length + myOutgoing.length}</p></div>
    </div>
  `;
}

async function showAdminPanel() {
  currentView = 'admin';
  updateActiveNavLink(4);
  const mainBody = document.getElementById('mainBody');
  const passwords = await getPasswords();
  const settings = await getSettings();

  mainBody.innerHTML = `
    <div class="animate-fade-in">
      <h1>Admin Control Panel</h1>
      <p style="color: var(--text-muted); margin-bottom: 2rem;">System Management & Global Security</p>

      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem;">
        <div class="card">
          <h3 style="margin-bottom: 1.5rem; display: flex; align-items: center; gap: 10px;">
            <i class="material-icons-round" style="color: var(--accent);">lock</i> Department Passwords
          </h3>
          <div style="max-height: 400px; overflow-y: auto;">
            ${departments.filter(d => d !== 'Administrator').map(d => `
              <div style="display: flex; justify-content: space-between; align-items: center; padding: 1rem 0; border-bottom: 1px solid #f1f5f9;">
                <div style="font-weight: 600;">${d}</div>
                <div style="display: flex; gap: 10px; align-items: center;">
                  <code style="background: #f1f5f9; padding: 4px 8px; border-radius: 4px;">${passwords[d] || '123456'}</code>
                  <button class="btn" style="padding: 4px 8px; font-size: 0.7rem;" onclick="adminResetPassword('${d}')">Reset</button>
                </div>
              </div>
            `).join('')}
          </div>
        </div>

        <div class="card">
          <h3 style="margin-bottom: 1.5rem; display: flex; align-items: center; gap: 10px;">
            <i class="material-icons-round" style="color: var(--accent);">history</i> Data Retention Settings
          </h3>
          <div style="padding: 1.5rem; background: #f8fafc; border-radius: 12px; margin-bottom: 1.5rem;">
            <label style="display: block; font-size: 0.8rem; font-weight: 700; color: #64748b; margin-bottom: 1rem;">AUTO-DELETE RECORDS OLDER THAN:</label>
            <select id="retentionSelect" class="form-control" style="width: 100%; padding: 0.8rem; border-radius: 8px; border: 1px solid #e2e8f0;" onchange="updateRetentionDays(this.value)">
              <option value="7" ${settings.retentionDays == 7 ? 'selected' : ''}>7 Days (Standard)</option>
              <option value="30" ${settings.retentionDays == 30 ? 'selected' : ''}>30 Days (Monthly)</option>
              <option value="90" ${settings.retentionDays == 90 ? 'selected' : ''}>90 Days (Quarterly)</option>
              <option value="-1" ${settings.retentionDays == -1 ? 'selected' : ''}>Keep Forever (Unlimited)</option>
            </select>
            <p style="font-size: 0.75rem; color: #94a3b8; margin-top: 1rem;">Note: Changes apply to all departments globally.</p>
          </div>
          
          <div style="padding: 1.5rem; background: #fff1f2; border-radius: 12px; border: 1px dashed #fda4af;">
            <h4 style="color: #e11d48; margin-bottom: 0.5rem;">Emergency Master Clear</h4>
            <p style="font-size: 0.8rem; color: #be123c; margin-bottom: 1rem;">This will permanently erase ALL letters and files from the Cloud database.</p>
            <button class="btn" style="background: #e11d48; color: white; width: 100%;" onclick="masterClear()">PURGE GLOBAL DATABASE</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

async function adminResetPassword(dept) {
  const newPass = prompt(`Enter new password for ${dept}:`, "123456");
  if (newPass) {
    const passwords = await getPasswords();
    passwords[dept] = newPass;
    await savePasswords(passwords);
    alert(`Password for ${dept} updated successfully!`);
    showAdminPanel();
  }
}

async function updateRetentionDays(days) {
  const settings = await getSettings();
  settings.retentionDays = parseInt(days);
  await saveSettings(settings);
  alert("Retention settings updated.");
}

async function masterClear() {
  if (confirm("🚨 WARNING: This will delete EVERYTHING in the cloud and local storage. Are you absolutely sure?")) {
    const letters = await getLetters();
    for (let l of letters) {
      await deleteLetter(l.id);
    }
    alert("System Cleared.");
    showDashboard();
  }
}

async function viewLetter(id) {
  const letters = await getLetters();
  const l = letters.find(lt => lt.id === id);
  if (!l) return;
  const isIncoming = l.recipient === currentDept;
  const isAdmin = currentDept === 'Administrator';

  const modalMarkup = `
    <div id="modalOverlay" style="position: fixed; inset: 0; background: rgba(15, 23, 42, 0.9); backdrop-filter: blur(10px); display: flex; align-items: center; justify-content: center; z-index: 2000;">
      <div class="card animate-fade-in" style="width: 100%; max-width: 700px; margin: 0 1rem; padding: 0; overflow: hidden; border: 1px solid rgba(255,255,255,0.1); box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);">
        <div style="background: linear-gradient(135deg, var(--navy), #1e293b); color: white; padding: 2rem; display: flex; justify-content: space-between; align-items: center;">
          <div>
            <h2 style="margin:0; font-size: 1.5rem; letter-spacing: 0.05em; color: var(--accent);">OFFICIAL RECORD</h2>
            <p style="margin: 0.25rem 0 0; font-size: 0.8rem; opacity: 0.7;">System Reference: <strong>${id}</strong></p>
          </div>
          <i class="material-icons-round" style="cursor: pointer; font-size: 2rem;" onclick="closeModal()">close</i>
        </div>
        
        <div style="padding: 2.5rem;">
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; margin-bottom: 2.5rem;">
            <div style="background: #f8fafc; padding: 1rem; border-radius: 8px;">
              <label style="display: block; font-size: 0.7rem; font-weight: 800; color: #64748b; text-transform: uppercase; margin-bottom: 0.5rem;">Originating Dept</label>
              <div style="font-weight: 600; color: var(--navy);">${l.sender}</div>
            </div>
            <div style="background: #f8fafc; padding: 1rem; border-radius: 8px;">
              <label style="display: block; font-size: 0.7rem; font-weight: 800; color: #64748b; text-transform: uppercase; margin-bottom: 0.5rem;">Recipient Dept</label>
              <div style="font-weight: 600; color: var(--navy);">${l.recipient}</div>
            </div>
          </div>

          ${l.isSummon ? `
            <div style="background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%); border: 2px solid #ef4444; border-radius: 14px; padding: 1.5rem; text-align: center; margin-bottom: 2rem; box-shadow: 0 4px 15px rgba(239, 68, 68, 0.15);">
              <div style="display: flex; align-items: center; justify-content: center; gap: 8px; color: #b91c1c; font-size: 1.15rem; font-weight: 800; margin-bottom: 0.5rem; text-transform: uppercase;">
                <i class="material-icons-round" style="font-size: 1.8rem; color: #ef4444;">campaign</i>
                Emergency Summon Notice
              </div>
              <div style="font-size: 1.35rem; font-weight: 800; color: #991b1b; margin: 0.5rem 0; line-height: 1.3;">
                "${l.message || 'Please come as soon as possible'}"
              </div>
              ${l.location ? `<div style="font-size: 0.95rem; color: #7f1d1d; font-weight: 600; margin-top: 0.5rem;"><i class="material-icons-round" style="font-size: 1rem; vertical-align: middle;">location_on</i> Report Location: <strong>${l.location}</strong></div>` : ''}
              ${(isIncoming || isAdmin) && l.status === 'urgent' ? `
                <button class="btn" style="background: #16a34a; color: white; font-weight: 700; margin-top: 1rem; padding: 0.75rem 2rem; border: none; border-radius: 8px; box-shadow: 0 4px 10px rgba(22, 163, 74, 0.35); cursor: pointer;" onclick="acknowledgeSummon('${l.id}')">
                  <i class="material-icons-round" style="font-size: 1.1rem; vertical-align: middle;">check_circle</i> I Am On My Way (Acknowledge)
                </button>
              ` : (l.status === 'acknowledged' ? `
                <div style="margin-top: 0.75rem; color: #166534; font-weight: 700; font-size: 0.85rem;"><i class="material-icons-round" style="vertical-align: middle; font-size: 1.1rem;">verified</i> Attendance has been acknowledged and confirmed.</div>
              ` : '')}
            </div>
          ` : ''}

          <div style="margin-bottom: 2.5rem;">
            <label style="display: block; font-size: 0.75rem; font-weight: 800; color: #64748b; text-transform: uppercase; margin-bottom: 0.5rem;">Letter Subject</label>
            <div style="font-size: 1.3rem; font-weight: 800; color: var(--navy); line-height: 1.3;">${l.subject}</div>
          </div>

          ${(l.fileBlob || l.fileUrl || (l.fileChunks && l.fileChunks > 0)) ? `
            <div style="background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%); border: 2px solid var(--accent); border-radius: 16px; padding: 2rem; text-align: center; margin-bottom: 2.5rem; position: relative; overflow: hidden;">
              <div style="position: absolute; top: 0; right: 0; padding: 0.5rem 1rem; background: var(--accent); color: white; font-size: 0.7rem; font-weight: 800; border-bottom-left-radius: 12px;">CLOUD DOCUMENT READY</div>
              <i class="material-icons-round" style="font-size: 4rem; color: var(--accent); margin-bottom: 1rem;">verified_user</i>
              <h3 style="margin-bottom: 0.5rem; color: var(--navy);">Access Authenticated Document</h3>
              <p style="color: #475569; font-size: 0.95rem; margin-bottom: 2rem;">Filename: <strong>${l.fileName || 'Document'}</strong></p>
              
              <button class="btn btn-primary" style="width: 100%; padding: 1.25rem; font-size: 1.2rem; justify-content: center; letter-spacing: 0.05em; background: #1e293b;" onclick="downloadFile('${id}')">
                <i class="material-icons-round" style="font-size: 1.5rem;">file_download</i> GET ORIGINAL FILE
              </button>
            </div>
          ` : `
            <div style="padding: 2.5rem; text-align: center; background: #fff1f2; border: 1px dashed #fda4af; border-radius: 16px; margin-bottom: 2.5rem;">
              <i class="material-icons-round" style="font-size: 3rem; color: #e11d48; margin-bottom: 1rem;">error_outline</i>
              <p style="color: #9f1239; font-weight: 700; font-size: 1.1rem;">No file attached.</p>
            </div>
          `}

          <div style="display: flex; justify-content: space-between; align-items: center; padding-top: 1.5rem; border-top: 1px solid #e2e8f0;">
            <div style="display: flex; align-items: center; gap: 10px;">
              <div style="font-size: 0.7rem; font-weight: 800; color: #64748b;">WORKFLOW:</div>
              <span style="padding: 0.3rem 0.8rem; border-radius: 99px; font-size: 0.75rem; font-weight: 800; background: #e2e8f0; color: #1e293b;">${l.status.toUpperCase()}</span>
            </div>
            <div style="display: flex; gap: 1rem;">
              ${isAdmin ? `<button class="btn" style="background: #e11d48; color: white;" onclick="adminRemoveLetter('${id}')">Master Delete</button>` : ''}
              <button class="btn" style="background: #f1f5f9; color: #475569;" onclick="printLetter('${id}')"><i class="material-icons-round" style="font-size: 1.1rem; vertical-align: middle;">print</i> Slip</button>
              ${isIncoming && l.status !== 'processed' ? `
                <button class="btn" style="background: #22c55e; color: white;" onclick="changeStatus('${id}', 'processed')">Finalize Action</button>
              ` : ''}
              <button class="btn" style="background: #1e293b; color: white;" onclick="closeModal()">Close</button>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', modalMarkup);
}

async function adminRemoveLetter(id) {
  if (confirm("Delete this record globally?")) {
    await deleteLetter(id);
    closeModal();
    showDashboard();
  }
}

async function changeStatus(id, newStatus) {
  await updateLetterStatus(id, newStatus);
  closeModal();
  showIncoming();
}

function closeModal() { const modal = document.getElementById('modalOverlay'); if (modal) modal.remove(); }

// Helper: Convert Base64 to Blob
const base64ToBlob = (base64, type = 'application/octet-stream') => {
  const binStr = atob(base64.split(',')[1] || base64);
  const len = binStr.length;
  const arr = new Uint8Array(len);
  for (let i = 0; i < len; i++) arr[i] = binStr.charCodeAt(i);
  return new Blob([arr], { type: type });
};

// Helper: Fetch Chunks & Save Blob (Silent)
async function fetchAndSaveBlob(l) {
  if (!l || !l.fileChunks) return null;
  console.log(`📥 Auto-Fetching ${l.fileChunks} chunks for ${l.id}...`);

  try {
    const snapshot = await db_cloud.collection('letters').doc(l.id).collection('chunks').orderBy('index').get();
    if (snapshot.empty) throw new Error("File parts missing in cloud.");

    let completeBase64 = '';
    snapshot.forEach(doc => { completeBase64 += doc.data().data; });

    console.log("✅ Auto-Assembly Complete.");
    const blob = base64ToBlob(completeBase64, l.fileType);

    // Save to Local DB
    l.fileBlob = blob;
    const tx = db_local.transaction(['letters'], 'readwrite');
    tx.objectStore('letters').put(l);

    return blob;
  } catch (e) {
    console.error("Auto-Fetch Failed:", e);
    return null;
  }
}

async function downloadFile(id) {
  const letters = await getLetters();
  let l = letters.find(lt => lt.id === id);
  if (!l) return;

  // 1. Try Local Blob (Fastest)
  if (l.fileBlob) {
    const url = URL.createObjectURL(l.fileBlob);
    triggerDownload(url, l.fileName || `letter_${id}`);
    return;
  }

  // 2. Try Chunked Cloud Download (No Card Mode)
  if (l.fileChunks > 0 && typeof db_cloud !== 'undefined') {
    const btn = document.querySelector(`button[onclick="downloadFile('${id}')"]`);
    const originalText = btn ? btn.innerHTML : '';
    if (btn) btn.innerHTML = `<i class='material-icons-round animate-spin'>cloud_download</i> Fetching Parts...`;

    // Use shared helper
    const blob = await fetchAndSaveBlob(l);

    if (blob) {
      if (btn) btn.innerHTML = `<i class='material-icons-round'>check</i> Ready`;
      const url = URL.createObjectURL(blob);
      triggerDownload(url, l.fileName);
      setTimeout(() => { if (btn) btn.innerHTML = originalText; }, 2000);
    } else {
      if (btn) btn.innerHTML = `<i class='material-icons-round'>error</i> Failed`;
      alert("Download Failed: Could not retrieve file parts.");
    }
  }
  // 3. Fallback to old URL method (legacy support)
  else if (l.fileUrl) {
    window.open(l.fileUrl, '_blank');
  }
}

function triggerDownload(url, name) {
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// --- EMERGENCY SUMMON MODULE ---
function playEmergencyTone() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'sawtooth';
    osc.frequency.setValueAtTime(880, ctx.currentTime);
    osc.frequency.setValueAtTime(659.25, ctx.currentTime + 0.15);
    osc.frequency.setValueAtTime(880, ctx.currentTime + 0.3);
    gain.gain.setValueAtTime(0.25, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.55);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.55);
  } catch (e) {
    console.warn("Audio alert unavailable:", e);
  }
}

const alertedSummons = new Set();

function showEmergencyAlertModal(summon) {
  if (alertedSummons.has(summon.id)) return;
  alertedSummons.add(summon.id);

  playEmergencyTone();

  const existing = document.getElementById('summonAlertOverlay');
  if (existing) existing.remove();

  const alertMarkup = `
    <div id="summonAlertOverlay" style="position: fixed; inset: 0; background: rgba(185, 28, 28, 0.88); backdrop-filter: blur(8px); display: flex; align-items: center; justify-content: center; z-index: 3000;">
      <div class="card animate-fade-in" style="width: 100%; max-width: 540px; margin: 0 1.5rem; text-align: center; border: 3px solid #ef4444; box-shadow: 0 25px 60px rgba(0,0,0,0.6); padding: 2.5rem 2rem;">
        <div style="width: 72px; height: 72px; background: #fee2e2; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 1.25rem; color: #ef4444;">
          <i class="material-icons-round" style="font-size: 2.8rem;">campaign</i>
        </div>
        <span style="background: #ef4444; color: white; padding: 4px 14px; border-radius: 20px; font-weight: 800; font-size: 0.75rem; letter-spacing: 0.05em; text-transform: uppercase;">🚨 EMERGENCY SUMMON (U YEERID DEG-DEG AH)</span>
        <h2 style="font-size: 1.55rem; margin: 1rem 0 0.5rem; color: #991b1b; font-weight: 800;">YOU ARE SUMMONED IMMEDIATELY!</h2>
        <p style="color: #64748b; font-size: 0.95rem; margin-bottom: 1.5rem;">Summoned by: <strong style="color: var(--navy); font-size: 1.05rem;">${summon.sender}</strong></p>
        
        <div style="background: #fff1f2; border: 2px dashed #f87171; border-radius: 12px; padding: 1.25rem; margin-bottom: 1.75rem;">
          <div style="font-size: 1.35rem; font-weight: 800; color: #b91c1c; margin-bottom: 0.5rem; line-height: 1.3;">
            "${summon.message || 'Please come as soon as possible'}"
          </div>
          ${summon.location ? `<div style="font-size: 0.9rem; color: #7f1d1d; font-weight: 600;"><i class="material-icons-round" style="font-size: 1rem; vertical-align: middle;">location_on</i> Report Location: <strong>${summon.location}</strong></div>` : ''}
        </div>

        <div style="display: flex; gap: 1rem; justify-content: center;">
          <button class="btn btn-primary" style="background: #16a34a; border: none; font-size: 1.05rem; padding: 0.85rem 2rem; font-weight: 700; display: inline-flex; align-items: center; gap: 8px; box-shadow: 0 4px 12px rgba(22, 163, 74, 0.4); border-radius: 8px; cursor: pointer;" onclick="acknowledgeSummon('${summon.id}')">
            <i class="material-icons-round">check_circle</i>
            I Am On My Way (Acknowledge)
          </button>
        </div>
      </div>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', alertMarkup);
}

async function acknowledgeSummon(id) {
  const overlay = document.getElementById('summonAlertOverlay');
  if (overlay) overlay.remove();
  closeModal();
  try {
    await updateLetterStatus(id, 'acknowledged', {
      acknowledgedAt: Date.now(),
      acknowledgedBy: currentDept
    });
  } catch(e) {
    console.warn("Could not update status:", e);
  }
  alert("✅ Acknowledged! The sender has been notified that you are on your way.");
  refreshCurrentView();
}

function playSuccessChime() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'sine';
    osc.frequency.setValueAtTime(523.25, ctx.currentTime); // C5
    osc.frequency.setValueAtTime(659.25, ctx.currentTime + 0.12); // E5
    osc.frequency.setValueAtTime(783.99, ctx.currentTime + 0.24); // G5
    gain.gain.setValueAtTime(0.25, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.65);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.65);
  } catch(e) {}
}

const alertedSenderSummons = new Set();

function showSenderSummonAckNotification(summon) {
  const key = `${summon.id}_ack`;
  if (alertedSenderSummons.has(key)) return;
  alertedSenderSummons.add(key);

  playSuccessChime();

  const toastId = `toast_${Date.now()}`;
  const toastMarkup = `
    <div id="${toastId}" class="animate-fade-in" style="position: fixed; top: 80px; right: 24px; background: #ffffff; border: 2.5px solid #22c55e; border-radius: 14px; box-shadow: 0 15px 35px rgba(0,0,0,0.2); padding: 1.25rem 1.5rem; z-index: 99999; display: flex; align-items: flex-start; gap: 1rem; width: 390px; transition: all 0.3s ease;">
      <div style="width: 46px; height: 46px; border-radius: 50%; background: #dcfce7; color: #16a34a; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
        <i class="material-icons-round" style="font-size: 2rem;">directions_run</i>
      </div>
      <div style="flex: 1;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.25rem;">
          <span style="font-size: 0.7rem; font-weight: 800; background: #22c55e; color: white; padding: 2px 8px; border-radius: 4px; text-transform: uppercase;">Summon Response</span>
          <i class="material-icons-round" style="cursor: pointer; color: #94a3b8; font-size: 1.2rem;" onclick="document.getElementById('${toastId}').remove()">close</i>
        </div>
        <div style="font-size: 1.05rem; font-weight: 800; color: #14532d; margin-bottom: 0.25rem;">
          ${summon.recipient} is on the way!
        </div>
        <div style="font-size: 0.85rem; color: #475569;">
          Confirmed receipt of your emergency summon: <em>"${summon.message || 'Please come as soon as possible'}"</em>
        </div>
        <div style="font-size: 0.75rem; color: #16a34a; font-weight: 700; margin-top: 0.4rem; display: flex; align-items: center; gap: 4px;">
          <i class="material-icons-round" style="font-size: 0.95rem;">verified</i> Acknowledged & Coming
        </div>
      </div>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', toastMarkup);

  updateSidebarNotifications();

  setTimeout(() => {
    const el = document.getElementById(toastId);
    if (el) el.remove();
  }, 12000);
}

async function toggleNotificationMenu() {
  const menu = document.getElementById('notifDropdown');
  if (!menu) return;

  if (menu.style.display === 'block') {
    menu.style.display = 'none';
    return;
  }

  const letters = await getLetters();
  
  // Summons sent by this department
  const mySentSummons = letters.filter(l => l.sender === currentDept && l.isSummon);
  // Summons received by this department
  const myIncomingSummons = letters.filter(l => (l.recipient === currentDept || l.recipient === 'ALL') && l.isSummon);
  // Recent incoming regular letters
  const myIncomingLetters = letters.filter(l => l.recipient === currentDept && !l.isSummon && l.status === 'new');

  menu.innerHTML = `
    <div style="padding: 1rem 1.25rem; background: var(--navy); color: white; display: flex; justify-content: space-between; align-items: center;">
      <div style="font-weight: 700; font-size: 0.95rem; display: flex; align-items: center; gap: 6px;">
        <i class="material-icons-round" style="color: var(--accent); font-size: 1.2rem;">notifications</i> Notifications & Summon Status
      </div>
      <span style="font-size: 0.75rem; background: rgba(255,255,255,0.15); padding: 2px 8px; border-radius: 12px;">Real-Time</span>
    </div>

    <div style="max-height: 420px; overflow-y: auto; padding: 0.5rem 0;">
      <!-- Section: Outgoing Summons Sent By You -->
      <div style="padding: 0.5rem 1rem; font-size: 0.75rem; font-weight: 800; color: #64748b; background: #f8fafc; text-transform: uppercase;">
        📡 Summons Dispatched By You (${mySentSummons.length})
      </div>
      ${mySentSummons.length === 0 ? `
        <div style="padding: 0.75rem 1rem; font-size: 0.85rem; color: #94a3b8; font-style: italic;">No emergency summons sent yet.</div>
      ` : mySentSummons.sort((a,b) => b.timestamp - a.timestamp).slice(0, 5).map(s => `
        <div style="padding: 0.75rem 1rem; border-bottom: 1px solid #f1f5f9; display: flex; align-items: flex-start; gap: 10px;">
          <div style="width: 32px; height: 32px; border-radius: 50%; background: ${s.status === 'acknowledged' ? '#dcfce7' : '#fee2e2'}; color: ${s.status === 'acknowledged' ? '#16a34a' : '#ef4444'}; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
            <i class="material-icons-round" style="font-size: 1.1rem;">${s.status === 'acknowledged' ? 'directions_run' : 'hourglass_top'}</i>
          </div>
          <div style="flex: 1;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
              <strong style="font-size: 0.85rem; color: #1e293b;">To: ${s.recipient}</strong>
              <span style="font-size: 0.7rem; font-weight: 700; padding: 2px 7px; border-radius: 4px; background: ${s.status === 'acknowledged' ? '#dcfce7' : '#fef3c7'}; color: ${s.status === 'acknowledged' ? '#166534' : '#92400e'};">
                ${s.status === 'acknowledged' ? '🟢 ON THE WAY' : '🟡 WAITING'}
              </span>
            </div>
            <div style="font-size: 0.8rem; color: #64748b; margin-top: 2px;">"${s.message || 'Please come as soon as possible'}"</div>
            <div style="font-size: 0.7rem; color: #94a3b8; margin-top: 2px;">${s.date} ${s.location ? `• 📍 ${s.location}` : ''}</div>
          </div>
        </div>
      `).join('')}

      <!-- Section: Incoming Alerts -->
      <div style="padding: 0.5rem 1rem; font-size: 0.75rem; font-weight: 800; color: #64748b; background: #f8fafc; text-transform: uppercase; margin-top: 0.5rem;">
        📥 Incoming Summons & New Letters
      </div>
      ${myIncomingSummons.filter(s => s.status === 'urgent').length === 0 && myIncomingLetters.length === 0 ? `
        <div style="padding: 0.75rem 1rem; font-size: 0.85rem; color: #94a3b8; font-style: italic;">No unread incoming notifications.</div>
      ` : `
        ${myIncomingSummons.filter(s => s.status === 'urgent').map(s => `
          <div style="padding: 0.75rem 1rem; border-bottom: 1px solid #fee2e2; background: #fff5f5; display: flex; align-items: flex-start; gap: 10px;">
            <i class="material-icons-round" style="color: #ef4444; font-size: 1.3rem;">campaign</i>
            <div style="flex: 1;">
              <strong style="color: #991b1b; font-size: 0.85rem;">URGENT SUMMON from ${s.sender}</strong>
              <div style="font-size: 0.8rem; color: #7f1d1d;">"${s.message || 'Please come as soon as possible'}"</div>
              <button class="btn" style="background: #16a34a; color: white; padding: 2px 8px; font-size: 0.7rem; font-weight: 700; margin-top: 4px; border: none; border-radius: 4px; cursor: pointer;" onclick="acknowledgeSummon('${s.id}')">Acknowledge</button>
            </div>
          </div>
        `).join('')}
        ${myIncomingLetters.map(l => `
          <div style="padding: 0.75rem 1rem; border-bottom: 1px solid #f1f5f9; display: flex; align-items: center; gap: 10px;">
            <i class="material-icons-round" style="color: #3b82f6; font-size: 1.2rem;">mark_email_unread</i>
            <div style="flex: 1;">
              <div style="font-size: 0.85rem; font-weight: 600;">From ${l.sender}: ${l.subject}</div>
              <div style="font-size: 0.7rem; color: #94a3b8;">${l.date}</div>
            </div>
          </div>
        `).join('')}
      `}
    </div>

    <div style="padding: 0.75rem; text-align: center; background: #f8fafc; border-top: 1px solid #e2e8f0;">
      <button class="btn" style="font-size: 0.8rem; padding: 4px 14px;" onclick="document.getElementById('notifDropdown').style.display='none'">Close</button>
    </div>
  `;

  menu.style.display = 'block';
}

function openSummonModal() {
  const otherDepts = departments.filter(d => d !== currentDept);
  const modalMarkup = `
    <div id="modalOverlay" style="position: fixed; inset: 0; background: rgba(15, 23, 42, 0.85); backdrop-filter: blur(6px); display: flex; align-items: center; justify-content: center; z-index: 2100;">
      <div class="card animate-fade-in" style="width: 100%; max-width: 520px; margin: 0 1rem; border: 2px solid #ef4444; box-shadow: 0 20px 40px -10px rgba(239, 68, 68, 0.4);">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px solid #fee2e2;">
          <div style="display: flex; align-items: center; gap: 10px;">
            <div style="background: #fee2e2; color: #ef4444; width: 44px; height: 44px; border-radius: 10px; display: flex; align-items: center; justify-content: center;">
              <i class="material-icons-round" style="font-size: 1.7rem;">campaign</i>
            </div>
            <div>
              <h2 style="margin: 0; font-size: 1.25rem; color: #991b1b; font-weight: 800;">Emergency Summon</h2>
              <span style="font-size: 0.75rem; color: #ef4444; font-weight: 700; text-transform: uppercase;">U Yeerid Deg-deg ah</span>
            </div>
          </div>
          <i class="material-icons-round" style="cursor: pointer; color: #64748b; font-size: 1.5rem;" onclick="closeModal()">close</i>
        </div>

        <form id="summonForm" onsubmit="handleSummonSubmit(event)">
          <div style="margin-bottom: 1.25rem;">
            <label style="display: block; font-size: 0.8rem; font-weight: 700; color: #1e293b; margin-bottom: 0.4rem;">
              Target Department <span style="color:#ef4444;">*</span>
            </label>
            <select id="summon_recipient" style="width: 100%; padding: 0.75rem; border-radius: 8px; border: 1.5px solid #cbd5e1; font-size: 0.95rem; font-weight: 500;" required>
              <option value="">-- Choose Department to Summon --</option>
              <option value="ALL" style="font-weight: 700; color: #b91c1c;">📢 ALL DEPARTMENTS (Broadcast Summon)</option>
              ${otherDepts.map(d => `<option value="${d}">${d}</option>`).join('')}
            </select>
          </div>

          <div style="margin-bottom: 1.25rem;">
            <label style="display: block; font-size: 0.8rem; font-weight: 700; color: #1e293b; margin-bottom: 0.4rem;">
              Location / Report To (Optional)
            </label>
            <input type="text" id="summon_location" value="Main Office / Xafiiska Wasiirka" placeholder="e.g. Main Office / Meeting Room" style="width: 100%; padding: 0.75rem; border-radius: 8px; border: 1.5px solid #cbd5e1; font-size: 0.95rem;">
          </div>

          <div style="margin-bottom: 1.5rem;">
            <label style="display: block; font-size: 0.8rem; font-weight: 700; color: #1e293b; margin-bottom: 0.4rem;">
              Summon Message <span style="color:#ef4444;">*</span>
            </label>
            <div style="background: #fef2f2; border: 1.5px solid #fecaca; border-radius: 8px; padding: 0.75rem;">
              <textarea id="summon_message" rows="3" required style="width: 100%; border: none; background: transparent; resize: none; font-size: 1rem; font-weight: 700; color: #991b1b; outline: none; font-family: inherit;">Please come as soon as possible</textarea>
              <div style="font-size: 0.75rem; color: #dc2626; margin-top: 4px; display: flex; align-items: center; gap: 4px;">
                <i class="material-icons-round" style="font-size: 0.9rem;">notification_important</i> An instant emergency sound and alert will trigger on the recipient's screen.
              </div>
            </div>
          </div>

          <div style="display: flex; justify-content: flex-end; gap: 0.75rem; padding-top: 1rem; border-top: 1px solid #f1f5f9;">
            <button type="button" class="btn" onclick="closeModal()" style="background: #f1f5f9; color: #475569; font-weight: 600;">Cancel</button>
            <button type="submit" class="btn" id="summonSubmitBtn" style="background: linear-gradient(135deg, #ef4444, #dc2626); color: white; border: none; font-weight: 700; padding: 0.75rem 1.5rem; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 4px 12px rgba(239, 68, 68, 0.4); border-radius: 8px; cursor: pointer;">
              <i class="material-icons-round" style="font-size: 1.2rem;">campaign</i>
              Dispatch Emergency Summon
            </button>
          </div>
        </form>
      </div>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', modalMarkup);
}

async function handleSummonSubmit(e) {
  e.preventDefault();
  const btn = document.getElementById('summonSubmitBtn');
  btn.disabled = true;
  btn.innerHTML = `<i class="material-icons-round animate-spin">refresh</i> Dispatching Summon...`;

  const recipient = document.getElementById('summon_recipient').value;
  const message = document.getElementById('summon_message').value.trim() || "Please come as soon as possible";
  const location = document.getElementById('summon_location').value.trim() || "Main Office";

  const targets = (recipient === 'ALL') ? departments.filter(d => d !== currentDept) : [recipient];

  try {
    for (const tgt of targets) {
      const summonRecord = {
        id: `SUMMON-${String(Date.now()).slice(-5)}-${Math.floor(Math.random() * 90 + 10)}`,
        sender: currentDept,
        recipient: tgt,
        subject: `🚨 EMERGENCY SUMMON: ${message}`,
        message: message,
        location: location,
        isSummon: true,
        priority: 'EMERGENCY',
        date: new Date().toISOString().split('T')[0],
        status: 'urgent',
        fileName: null,
        fileBlob: null,
        timestamp: Date.now()
      };
      await saveLetter(summonRecord);
    }
    alert(`🚨 Emergency Summon dispatched successfully to ${recipient === 'ALL' ? 'ALL Departments' : recipient}!`);
    closeModal();
    showDashboard();
  } catch (err) {
    console.error("Summon send error:", err);
    alert("❌ Error sending summon: " + err.message);
    btn.disabled = false;
    btn.innerHTML = `<i class="material-icons-round">campaign</i> Retry Summon`;
  }
}

function openNewLetterModal(defaultType = 'outgoing') {
  const otherDepts = departments.filter(d => d !== currentDept);
  const modalMarkup = `
    <div id="modalOverlay" style="position: fixed; inset: 0; background: rgba(0,0,0,0.5); backdrop-filter: blur(4px); display: flex; align-items: center; justify-content: center; z-index: 1000;">
      <div class="card animate-fade-in" style="width: 100%; max-width: 500px; margin: 0 1rem;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1.5rem;"><h2>Send New Letter</h2><i class="material-icons-round" style="cursor: pointer;" onclick="closeModal()">close</i></div>
        <form id="letterForm" onsubmit="handleLetterSubmit(event)">
          <div style="margin-bottom: 1rem;"><label>Recipient</label><select id="l_recipient" style="width:100%;padding:0.6rem;" required><option value="">Select...</option>${otherDepts.map(d => `<option value="${d}">${d}</option>`).join('')}</select></div>
          <div style="margin-bottom: 1rem;"><label>Subject</label><input type="text" id="l_subject" required style="width:100%;padding:0.6rem;"></div>
          <div style="margin-bottom: 1rem; border: 2px dashed #e2e8f0; padding:1.5rem; text-align:center;"><input type="file" id="l_file" style="display:none;" onchange="updateFileLabel(this)"><button type="button" class="btn" onclick="document.getElementById('l_file').click()">Select Original File</button><div id="fileNameDisplay" style="margin-top:0.5rem; font-size:0.8rem;"></div></div>
          <div style="display:flex; justify-content:flex-end; gap:1rem;"><button type="button" class="btn" onclick="closeModal()">Cancel</button><button type="submit" class="btn btn-primary" id="dispatchBtn">Send</button></div>
        </form>
      </div>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', modalMarkup);
}

function updateFileLabel(input) { if (input.files[0]) document.getElementById('fileNameDisplay').innerText = input.files[0].name; }

async function handleLetterSubmit(e) {
  e.preventDefault();
  const dispatchBtn = document.getElementById('dispatchBtn');
  dispatchBtn.disabled = true;
  dispatchBtn.innerText = "Sending to Cloud..."; // Specific feedback

  const fileInput = document.getElementById('l_file');
  const letter = {
    id: `LMS-${String(Date.now()).slice(-6)}`,
    sender: currentDept,
    recipient: document.getElementById('l_recipient').value,
    subject: document.getElementById('l_subject').value,
    date: new Date().toISOString().split('T')[0],
    status: 'new',
    fileBlob: fileInput.files[0] || null,
    fileName: fileInput.files[0] ? fileInput.files[0].name : null,
    timestamp: Date.now()
  };

  try {
    await saveLetter(letter); // Waits for Cloud confirmation
    alert("✅ Sent Successfully!");
    closeModal();
    showDashboard();
  } catch (error) {
    // If we are here, saveLetter threw an error (Cloud failed)
    // Detailed error already alerted in saveLetter, so just reset UI
    dispatchBtn.disabled = false;
    dispatchBtn.innerText = "Retry Send";
  }
}

async function login() {
  const btn = document.querySelector('.login-container .btn-primary');
  if (btn) btn.innerText = "Verifying...";

  try {
    const dept = document.getElementById('login_dept').value;
    const pass = document.getElementById('login_pass').value;
    if (!dept) { alert("Please select a department"); if (btn) btn.innerText = "Enter System"; return; }

    const passwords = await getPasswords();
    // Fallback to default if not set in DB yet
    const storedPass = passwords[dept] || DEFAULT_PASS;

    if (storedPass === pass) {
      currentDept = dept;
      sessionStorage.setItem('lms_current_dept', dept);
      const overlay = document.getElementById('loginOverlay');
      if (overlay) overlay.remove();

      const display = document.getElementById('deptNameDisplay');
      if (display) display.innerText = dept;

      setupRealtimeListener(); // Start syncing immediately
      updateSidebarLinks();
      showDashboard();
    } else {
      alert('Invalid Password');
      if (btn) btn.innerText = "Enter System";
    }
  } catch (e) {
    console.error("Login Error:", e);
    alert("System Error: " + e.message);
    if (btn) btn.innerText = "Enter System";
  }
}



function logout() {
  console.log("🔒 Logging out user...");
  try {
    if (typeof unsubscribeLetters === 'function') unsubscribeLetters();
  } catch (e) {
    console.warn("Could not unsubscribe letters:", e);
  }
  try {
    if (typeof unsubscribeSettings === 'function') unsubscribeSettings();
  } catch (e) {
    console.warn("Could not unsubscribe settings:", e);
  }

  unsubscribeLetters = null;
  unsubscribeSettings = null;

  // Clear session & local storage
  sessionStorage.removeItem('lms_current_dept');
  sessionStorage.clear();
  localStorage.removeItem('lms_current_dept');
  currentDept = null;

  // Clear UI
  const display = document.getElementById('deptNameDisplay');
  if (display) display.innerText = '';
  const mainBody = document.getElementById('mainBody');
  if (mainBody) mainBody.innerHTML = '';
  const adminLink = document.getElementById('adminLink');
  if (adminLink) adminLink.remove();

  // Always reload window to ensure clean reset and show login screen
  window.location.reload();
}

async function checkAuth() {
  if (!currentDept) {
    showLoginScreen();
  } else {
    const display = document.getElementById('deptNameDisplay');
    if (display) display.innerText = currentDept;
    setupRealtimeListener(); // Ensure listener is active on reload
    updateSidebarLinks();
    showDashboard();
  }
}

function updateSidebarLinks() {
  const sidebarNav = document.querySelector('.nav-links');
  if (!sidebarNav) return;
  const existingAdminLink = document.getElementById('adminLink');
  if (currentDept === 'Administrator') {
    if (!existingAdminLink) {
      const adminLink = document.createElement('a');
      adminLink.href = "#";
      adminLink.id = "adminLink";
      adminLink.className = "nav-item";
      adminLink.onclick = showAdminPanel;
      adminLink.innerHTML = `<i class="material-icons-round">settings_suggest</i><span>Admin Panel</span>`;
      sidebarNav.appendChild(adminLink);
    }
  } else if (existingAdminLink) {
    existingAdminLink.remove();
  }
}

async function showLoginScreen() {
  // Remove any existing overlay to prevent duplicates
  const existingOverlay = document.getElementById('loginOverlay');
  if (existingOverlay) existingOverlay.remove();

  const loginHtml = `
    <div id="loginOverlay" style="position:fixed; inset:0; background:var(--bg-sidebar); display:flex; align-items:center; justify-content:center; z-index:99999;">
      <div class="card animate-fade-in" style="width:400px; text-align:center; padding: 2rem;">
        <img src="logo.png" alt="Office Logo" style="width: 80px; height: 80px; object-fit: contain; margin:0 auto 1rem; border-radius: 12px;" onerror="this.style.display='none'">
        <h2 style="font-size: 1.15rem; line-height: 1.4; margin-bottom: 0.35rem; color: var(--primary);">Somali Revenue Bureau DDS</h2>
        <p style="color: var(--text-muted); font-size: 0.85rem; margin-bottom: 1.5rem;">Letter Management System</p>
        <div style="text-align:left; margin-top:1.5rem;">
          <label style="font-size: 0.8rem; font-weight: 600; color: #475569; display: block; margin-bottom: 0.35rem;">Department</label>
          <select id="login_dept" style="width:100%; padding:0.75rem; margin-bottom:1rem; border-radius: var(--radius-md); border: 1px solid #cbd5e1;">
            <option value="">Select Department...</option>
            ${departments.map(d => `<option value="${d}">${d}</option>`).join('')}
          </select>
          <label style="font-size: 0.8rem; font-weight: 600; color: #475569; display: block; margin-bottom: 0.35rem;">Password</label>
          <input type="password" id="login_pass" placeholder="Password (default: 123456)" style="width:100%; padding:0.75rem; border-radius: var(--radius-md); border: 1px solid #cbd5e1;" onkeypress="if(event.key==='Enter') login()">
        </div>
        <button class="btn btn-primary" style="width:100%; margin-top:1.5rem; justify-content: center; padding: 0.75rem;" onclick="login()">Enter System</button>
      </div>
    </div>`;
  document.body.insertAdjacentHTML('beforeend', loginHtml);
  setTimeout(() => {
    const deptSelect = document.getElementById('login_dept');
    if (deptSelect) deptSelect.focus();
  }, 100);
}

async function printLetter(id) {
  const letters = await getLetters();
  const l = letters.find(lt => lt.id === id);
  const w = window.open('', '_blank');
  w.document.write(`<html><head><title>Registry Slip</title><style>body{font-family:sans-serif;padding:40px;}</style></head><body><h1>REGISTRY SUMMARY SLIP</h1><p>ID: ${l.id}</p><p>From: ${l.sender}</p><p>To: ${l.recipient}</p><p>Subject: ${l.subject}</p><p>Date: ${l.date}</p><hr><p>Digitally registered in LMS Elite.</p></body></html>`);
  w.document.close(); w.print();
}

function openChangePasswordModal() {
  const h = `<div id="modalOverlay" style="position:fixed; inset:0; background:rgba(0,0,0,0.5); display:flex; align-items:center; justify-content:center; z-index:2000;"><div class="card"><h2>New Password</h2><input type="password" id="new_pass" style="width:100%;margin-top:1rem;padding:0.75rem;"><button class="btn btn-primary" onclick="submitNewPassword()" style="margin-top:1rem;width:100%;">Update</button><button class="btn" onclick="closeModal()" style="margin-top:0.5rem;width:100%;">Cancel</button></div></div>`;
  document.body.insertAdjacentHTML('beforeend', h);
}

async function submitNewPassword() {
  const passwords = await getPasswords();
  passwords[currentDept] = document.getElementById('new_pass').value;
  await savePasswords(passwords);
  closeModal();


  alert('Password Updated Successfully');
}

// Global Sync Listener (Restored Single Source of Truth)
let unsubscribeLetters = null;
let unsubscribeSettings = null;

function setupRealtimeListener() {
  if (typeof db_cloud === 'undefined' || !db_cloud) return;
  if (unsubscribeLetters) unsubscribeLetters();

  console.log("📡 Starting Global Cloud Listener (v18)...");
  unsubscribeLetters = db_cloud.collection('letters').onSnapshot(snapshot => {
    const transaction = db_local.transaction(['letters'], 'readwrite');
    const store = transaction.objectStore('letters');

    snapshot.docChanges().forEach(change => {
      const cloudData = change.doc.data();
      if (change.type === "added" || change.type === "modified") {
        // SMART MERGE: Preserve local FileBlob if it exists
        const request = store.get(change.doc.id);
        request.onsuccess = (event) => {
          const localData = event.target.result;
          if (localData && localData.fileBlob) {
            cloudData.fileBlob = localData.fileBlob;
          }
          store.put(cloudData);

          // AUTO-FETCH TRIGGER (User Request)
          // If I am the recipient, and I don't have the blob yet, and it has chunks... get it!
          if (currentDept && cloudData.recipient === currentDept && !cloudData.fileBlob && cloudData.fileChunks > 0) {
            fetchAndSaveBlob(cloudData).then(blob => {
              if (blob) {
                console.log("⚡ Auto-Fetch Success for", cloudData.id);
                refreshCurrentView(); // Refresh to show "Open File" button state
              }
            });
          }

          // Real-time Emergency Summon Alert Check (For Recipient)
          if (currentDept && cloudData.isSummon && (cloudData.recipient === currentDept || cloudData.recipient === 'ALL') && cloudData.status === 'urgent') {
            showEmergencyAlertModal(cloudData);
          }

          // Real-time Summon Acknowledgment Alert (For SENDER who sent the summon)
          if (currentDept && cloudData.isSummon && cloudData.sender === currentDept && cloudData.status === 'acknowledged') {
            showSenderSummonAckNotification(cloudData);
          }
        };
        request.onerror = () => store.put(cloudData);
      }
      if (change.type === "removed") {
        store.delete(change.doc.id);
      }
    });

    transaction.oncomplete = () => {
      console.log("☁️ Cloud Update Synced to Local!");
      refreshCurrentView();
      updateSidebarLinks(); // Notification badges
    };
  });

  // Settings Listener
  if (unsubscribeSettings) unsubscribeSettings();
  unsubscribeSettings = db_cloud.collection('settings').onSnapshot(snapshot => {
    snapshot.docChanges().forEach(change => {
      if (change.doc.id === 'passwords') {
        const tx = db_local.transaction(['settings'], 'readwrite');
        tx.objectStore('settings').put(change.doc.data(), 'passwords');
        console.log("🔐 Passwords Synced");
      }
    });
  });
}

let currentView = 'dashboard'; // tracks what user is looking at

function refreshCurrentView() {
  if (!currentDept) return;
  if (currentView === 'dashboard') showDashboard();
  else if (currentView === 'incoming') showIncoming();
  else if (currentView === 'outgoing') showOutgoing();
  else if (currentView === 'reports') showReports();
  else if (currentView === 'admin') showAdminPanel();
}

// Initialization
document.addEventListener('DOMContentLoaded', async () => {
  await initDB();
  try { await cleanupExpiredLetters(); } catch (e) { console.warn("Cleanup Error:", e); }

  // Auth check will start listener if logged in
  checkAuth();

  const searchInput = document.querySelector('.header input');
  if (searchInput) { searchInput.addEventListener('keypress', (e) => { if (e.key === 'Enter') handleSearch(searchInput.value); }); }

  const logoutBtn = document.querySelector('button[onclick="logout()"]');
  if (logoutBtn) {
    logoutBtn.addEventListener('click', (e) => {
      e.preventDefault();
      logout();
    });
  }
});
