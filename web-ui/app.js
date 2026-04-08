const state = {
  connected: false,
  busy: false
};

const elements = {
  connectionUri: document.getElementById("connectionUri"),
  authentication: document.getElementById("authentication"),
  username: document.getElementById("username"),
  password: document.getElementById("password"),
  startTime: document.getElementById("startTime"),
  endTime: document.getElementById("endTime"),
  fromField: document.getElementById("fromField"),
  subjectField: document.getElementById("subjectField"),
  keywordField: document.getElementById("keywordField"),
  targetMailbox: document.getElementById("targetMailbox"),
  targetFolder: document.getElementById("targetFolder"),
  usersField: document.getElementById("usersField"),
  reviewConfirmed: document.getElementById("reviewConfirmed"),
  connectButton: document.getElementById("connectButton"),
  disconnectButton: document.getElementById("disconnectButton"),
  loadUsersButton: document.getElementById("loadUsersButton"),
  saveUsersButton: document.getElementById("saveUsersButton"),
  cleanUsersButton: document.getElementById("cleanUsersButton"),
  searchButton: document.getElementById("searchButton"),
  deleteButton: document.getElementById("deleteButton"),
  queryPreview: document.getElementById("queryPreview"),
  commandPreview: document.getElementById("commandPreview"),
  resultTable: document.getElementById("resultTable"),
  resultSummary: document.getElementById("resultSummary"),
  logList: document.getElementById("logList"),
  connectionStatus: document.getElementById("connectionStatus"),
  sessionSummary: document.getElementById("sessionSummary"),
  statusDot: document.getElementById("statusDot")
};

function formatUtc(value) {
  if (!value) {
    return "";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()} ${date.getUTCHours()}:${date.getUTCMinutes()}:${date.getUTCSeconds()}`;
}

function normalizeUsersText(input) {
  return String(input || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.replace(/@brbiotech\.com$/i, ""))
    .join("\n");
}

function buildQueryPreview() {
  const parts = [];
  const start = formatUtc(elements.startTime.value);
  const end = formatUtc(elements.endTime.value);
  const from = elements.fromField.value.trim();
  const subject = elements.subjectField.value.trim();
  const keyword = elements.keywordField.value.trim();

  if (start || end) {
    parts.push(`Sent:"${start} ... ${end}"`);
  }
  if (from) {
    parts.push(`From:"${from}"`);
  }
  if (subject) {
    parts.push(`Subject:"${subject}"`);
  }
  if (keyword) {
    parts.push(`"${keyword}"`);
  }

  return parts.join(" AND ") || "Fill in a time range or at least one filter.";
}

function buildCommandPreview(mode) {
  const users = normalizeUsersText(elements.usersField.value).split("\n").filter(Boolean);
  if (!users.length) {
    return "# Fill in at least one target account";
  }

  const query = buildQueryPreview();
  const sampleUser = users[0];
  if (mode === "delete") {
    return `Search-Mailbox -SearchQuery '${query}' -Identity ${sampleUser} -DeleteContent -Force`;
  }

  return `Search-Mailbox -SearchQuery '${query}' -Identity ${sampleUser} -TargetMailbox ${elements.targetMailbox.value.trim() || "your.name@company.com"} -TargetFolder "${elements.targetFolder.value.trim() || "Deleted Items"}"`;
}

function getPayload() {
  return {
    startTime: elements.startTime.value,
    endTime: elements.endTime.value,
    from: elements.fromField.value.trim(),
    subject: elements.subjectField.value.trim(),
    keyword: elements.keywordField.value.trim(),
    targetMailbox: elements.targetMailbox.value.trim(),
    targetFolder: elements.targetFolder.value.trim(),
    users: elements.usersField.value
  };
}

async function api(url, options = {}) {
  const response = await fetch(url, {
    headers: {
      "Content-Type": "application/json"
    },
    ...options
  });
  const data = await response.json();
  if (!response.ok || !data.ok) {
    throw new Error(data.message || "Request failed.");
  }
  return data;
}

function setBusy(nextValue) {
  state.busy = nextValue;
  [
    elements.connectButton,
    elements.disconnectButton,
    elements.loadUsersButton,
    elements.saveUsersButton,
    elements.cleanUsersButton,
    elements.searchButton,
    elements.deleteButton
  ].forEach((button) => {
    button.disabled = nextValue;
  });
}

function addLog(message) {
  const time = new Date().toLocaleString("zh-CN", { hour12: false });
  const item = document.createElement("div");
  item.className = "log-item";
  item.innerHTML = `<strong>${time}</strong><br>${escapeHtml(message)}`;

  if (elements.logList.children.length === 1 && elements.logList.textContent.includes("Waiting for the next action")) {
    elements.logList.innerHTML = "";
  }

  elements.logList.prepend(item);
}

function escapeHtml(input) {
  return String(input || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function updatePreviews() {
  elements.queryPreview.textContent = buildQueryPreview();
  elements.commandPreview.textContent = buildCommandPreview("search");
}

function renderConnection(connected, session) {
  state.connected = Boolean(connected);
  elements.statusDot.classList.toggle("online", state.connected);
  elements.connectionStatus.textContent = state.connected ? "Connected" : "Disconnected";

  if (session) {
    elements.sessionSummary.textContent = `${session.username} | ${session.authentication} | ${session.connectionUri}`;
  } else {
    elements.sessionSummary.textContent = "Fill in the connection settings and validate the Exchange session first.";
  }
}

function renderResults(results, summary) {
  elements.resultSummary.textContent = summary;

  if (!results || !results.length) {
    elements.resultTable.innerHTML = '<tr><td colspan="5" class="empty-cell">No per-account result was returned.</td></tr>';
    return;
  }

  elements.resultTable.innerHTML = results
    .map((item) => {
      const actionText = item.actionTaken ? "Executed" : "Skipped";
      const statusClass = item.status === "success" ? "success" : "error";
      const statusText = item.status === "success" ? "Success" : "Error";
      return `
        <tr>
          <td>${escapeHtml(item.user)}</td>
          <td>${escapeHtml(String(item.estimatedCount ?? 0))}</td>
          <td>${escapeHtml(actionText)}</td>
          <td><span class="pill ${statusClass}">${escapeHtml(statusText)}</span></td>
          <td>${escapeHtml(item.message || "")}</td>
        </tr>
      `;
    })
    .join("");
}

async function loadState() {
  const data = await api("/api/state", { method: "GET" });
  if (data.defaults) {
    elements.connectionUri.value = data.defaults.connectionUri || "";
    elements.authentication.value = data.defaults.authentication || "Kerberos";
    elements.username.value = data.defaults.username || "";
    elements.targetFolder.value = data.defaults.targetFolder || "Deleted Items";
  }
  renderConnection(data.connected, data.session);
}

async function loadUsers() {
  const data = await api("/api/users", { method: "GET" });
  elements.usersField.value = data.usersText || "";
  updatePreviews();
}

async function handleConnect() {
  setBusy(true);
  try {
    const data = await api("/api/connect", {
      method: "POST",
      body: JSON.stringify({
        connectionUri: elements.connectionUri.value.trim(),
        authentication: elements.authentication.value,
        username: elements.username.value.trim(),
        password: elements.password.value
      })
    });
    renderConnection(true, data.session);
    elements.password.value = "";
    addLog(data.message);
  } catch (error) {
    addLog(`Connect failed: ${error.message}`);
    alert(error.message);
  } finally {
    setBusy(false);
  }
}

async function handleDisconnect() {
  setBusy(true);
  try {
    const data = await api("/api/disconnect", {
      method: "POST",
      body: JSON.stringify({})
    });
    renderConnection(false, null);
    elements.reviewConfirmed.checked = false;
    addLog(data.message);
  } catch (error) {
    addLog(`Disconnect failed: ${error.message}`);
    alert(error.message);
  } finally {
    setBusy(false);
  }
}

async function handleSaveUsers() {
  setBusy(true);
  try {
    const normalized = normalizeUsersText(elements.usersField.value);
    const data = await api("/api/users", {
      method: "POST",
      body: JSON.stringify({ users: normalized })
    });
    elements.usersField.value = data.usersText || normalized;
    updatePreviews();
    addLog(data.message);
  } catch (error) {
    addLog(`Save failed: ${error.message}`);
    alert(error.message);
  } finally {
    setBusy(false);
  }
}

function handleCleanUsers() {
  elements.usersField.value = normalizeUsersText(elements.usersField.value);
  updatePreviews();
  addLog("Mail suffix cleanup completed.");
}

async function handleSearch() {
  setBusy(true);
  try {
    const data = await api("/api/search", {
      method: "POST",
      body: JSON.stringify(getPayload())
    });
    renderResults(data.results, `${data.message} Query: ${data.query}`);
    elements.reviewConfirmed.checked = false;
    addLog(`${data.message} Review the copied mail before delete.`);
  } catch (error) {
    addLog(`Search failed: ${error.message}`);
    alert(error.message);
  } finally {
    setBusy(false);
  }
}

async function handleDelete() {
  if (!elements.reviewConfirmed.checked) {
    alert("Review the copied search result first, then confirm delete.");
    return;
  }

  const answer = window.confirm("Delete will run with the current filters. Make sure the matching search result has already been reviewed.");
  if (!answer) {
    return;
  }

  setBusy(true);
  try {
    const data = await api("/api/delete", {
      method: "POST",
      body: JSON.stringify({
        ...getPayload(),
        reviewConfirmed: true
      })
    });
    renderResults(data.results, `${data.message} Query: ${data.query}`);
    addLog(data.message);
  } catch (error) {
    addLog(`Delete failed: ${error.message}`);
    alert(error.message);
  } finally {
    setBusy(false);
  }
}

function bindEvents() {
  [
    elements.startTime,
    elements.endTime,
    elements.fromField,
    elements.subjectField,
    elements.keywordField,
    elements.targetMailbox,
    elements.targetFolder,
    elements.usersField
  ].forEach((element) => {
    element.addEventListener("input", updatePreviews);
  });

  elements.connectButton.addEventListener("click", handleConnect);
  elements.disconnectButton.addEventListener("click", handleDisconnect);
  elements.loadUsersButton.addEventListener("click", async () => {
    setBusy(true);
    try {
      await loadUsers();
      addLog("Loaded account list from leaverlist.ps1.");
    } catch (error) {
      addLog(`Load failed: ${error.message}`);
      alert(error.message);
    } finally {
      setBusy(false);
    }
  });
  elements.saveUsersButton.addEventListener("click", handleSaveUsers);
  elements.cleanUsersButton.addEventListener("click", handleCleanUsers);
  elements.searchButton.addEventListener("click", handleSearch);
  elements.deleteButton.addEventListener("click", handleDelete);
}

async function init() {
  bindEvents();
  updatePreviews();
  try {
    await loadState();
    await loadUsers();
    addLog("Page initialization completed.");
  } catch (error) {
    addLog(`Initialization failed: ${error.message}`);
  }
}

init();
