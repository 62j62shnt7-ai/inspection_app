const state = {
  view: "dashboard",
  assets: [],
  sheets: [],
  editingAsset: null,
  activeDrawerTab: "specs",
};

// ---------------------------------------------------------------- error toast
function showToast(msg, isError = true) {
  let container = document.getElementById("toastContainer");
  if (!container) {
    container = document.createElement("div");
    container.id = "toastContainer";
    document.body.appendChild(container);
  }
  const el = document.createElement("div");
  el.className = "toast" + (isError ? " error" : "");
  el.textContent = msg;
  container.appendChild(el);
  requestAnimationFrame(() => el.classList.add("show"));
  setTimeout(() => {
    el.classList.remove("show");
    setTimeout(() => el.remove(), 300);
  }, 5000);
}

// ---------------------------------------------------------------- helpers
function fmtDate(d) {
  if (!d) return "—";
  return d;
}

function daysFromToday(iso) {
  if (!iso) return null;
  const today = new Date();
  today.setHours(0,0,0,0);
  const target = new Date(iso + "T00:00:00");
  return Math.round((target - today) / 86400000);
}

function statusPill(asset) {
  const nd = asset.next_due;
  if (!nd) return `<span class="status-pill none">No due date</span>`;
  const days = daysFromToday(nd);
  if (days < 0) return `<span class="status-pill overdue">OVERDUE ${Math.abs(days)}d</span>`;
  if (days <= 30) return `<span class="status-pill soon">DUE ${days}d</span>`;
  return `<span class="status-pill ok">${fmtDate(nd)}</span>`;
}

function riskBadge(asset) {
  const r = (asset.risk_category || "LOW").toUpperCase();
  const cls = r === "HIGH" ? "high" : (r === "MEDIUM" ? "medium" : "low");
  const cuiHtml = asset.is_cui ? `<span class="badge-cui" title="Corrosion Under Insulation Susceptible">CUI</span>` : "";
  return `<span class="badge-risk ${cls}">${r}</span>${cuiHtml}`;
}

function esc(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/[&<>"']/g, c => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  })[c]);
}

async function api(path, opts) {
  const res = await fetch(path, opts);
  if (!res.ok) {
    const err = await res.json().catch(() => ({error: res.statusText}));
    throw new Error(err.error || "Request failed");
  }
  return res.json();
}

// ---------------------------------------------------------------- tabs
document.querySelectorAll(".tab").forEach(btn => {
  btn.addEventListener("click", () => switchView(btn.dataset.view));
});

function switchView(view) {
  state.view = view;
  document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b.dataset.view === view));
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  document.getElementById("view-" + view).classList.add("active");
  if (view === "dashboard") loadDashboard();
  if (view === "assets") loadAssets();
  if (view === "yearly") {
    if (!document.getElementById("yearInput").value) {
      document.getElementById("yearInput").value = new Date().getFullYear();
    }
  }
}

// ---------------------------------------------------------------- dashboard
function assetRowHtml(a, compact) {
  const tag = a.tag || a.sn || "";
  const loc = [a.plant, a.location].filter(Boolean).join(" · ") || "—";
  return `<tr data-id="${a.id}">
    <td><div>${esc(a.name || "(unnamed)")} ${riskBadge(a)}</div><div class="tag-mono">${esc(tag)}</div></td>
    <td>${esc(a.source_sheet || "")}</td>
    <td>${statusPill(a)}</td>
    ${compact ? "" : `<td><span class="tag-mono" style="font-weight:600;">${esc(a.field || "—")}</span></td><td>${esc(loc)}</td>`}
  </tr>`;
}

function tableHtml(rows, compact) {
  if (!rows.length) {
    return `<table class="assets"><tbody><tr class="empty-row"><td>No matching assets.</td></tr></tbody></table>`;
  }
  const extraHeads = compact ? "" : `<th>Pack # / Header</th><th>Plant & Location</th>`;
  return `<table class="assets">
    <thead><tr><th>Asset & Risk</th><th>Source</th><th>Status</th>${extraHeads}</tr></thead>
    <tbody>${rows.map(a => assetRowHtml(a, compact)).join("")}</tbody>
  </table>`;
}

function bindRowClicks(container) {
  container.querySelectorAll("tr[data-id]").forEach(tr => {
    tr.addEventListener("click", () => openAssetDrawer(parseInt(tr.dataset.id)));
  });
}

async function loadDashboard() {
  try {
    const d = await api("/api/dashboard");
    document.getElementById("statRow").innerHTML = `
      <div class="stat-card"><div class="n">${d.total_assets}</div><div class="label">Total tracked assets</div></div>
      <div class="stat-card overdue"><div class="n">${d.overdue_count}</div><div class="label">Overdue</div></div>
      <div class="stat-card overdue" style="background:#3a1f26;"><div class="n" style="color:#f56c6c;">${d.high_risk_count}</div><div class="label">High Risk (API 580)</div></div>
      <div class="stat-card" style="background:#2b1b36;"><div class="n" style="color:#d670f5;">${d.cui_count}</div><div class="label">CUI Flagged</div></div>
      <div class="stat-card due30"><div class="n">${d.due_30_count}</div><div class="label">Due in 30 days</div></div>
    `;

    const highRiskEl = document.getElementById("highRiskList");
    if (highRiskEl) {
      highRiskEl.innerHTML = tableHtml(d.high_risk || [], true);
      bindRowClicks(highRiskEl);
    }

    const overdueEl = document.getElementById("overdueList");
    overdueEl.innerHTML = tableHtml(d.overdue, true);
    bindRowClicks(overdueEl);

    const due30El = document.getElementById("due30List");
    due30El.innerHTML = tableHtml(d.due_30, true);
    bindRowClicks(due30El);

    const bySheetEl = document.getElementById("bySheet");
    bySheetEl.innerHTML = d.sheets.map(s => {
      const info = d.by_sheet[s];
      return `<span class="chip"><b>${esc(s)}</b> — ${info.total} ${info.overdue ? `· <span class="od">${info.overdue} overdue</span>` : ""}</span>`;
    }).join("");

    const sel = document.getElementById("sheetFilter");
    const current = sel.value;
    sel.innerHTML = `<option value="">All sheets</option>` + d.sheets.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
    sel.value = current;
  } catch (err) {
    showToast("Failed to load dashboard: " + err.message);
  }
}

// ---------------------------------------------------------------- assets list
let searchDebounce = null;
document.getElementById("searchBox").addEventListener("input", () => {
  clearTimeout(searchDebounce);
  searchDebounce = setTimeout(loadAssets, 200);
});
document.getElementById("sheetFilter").addEventListener("change", loadAssets);
document.getElementById("riskFilter")?.addEventListener("change", loadAssets);
document.getElementById("groupFilter")?.addEventListener("change", loadAssets);
document.getElementById("overdueOnly").addEventListener("change", loadAssets);

async function loadAssets() {
  try {
    const q = document.getElementById("searchBox").value;
    const sheet = document.getElementById("sheetFilter").value;
    const risk = document.getElementById("riskFilter")?.value || "";
    const overdue = document.getElementById("overdueOnly").checked ? "1" : "0";
    const params = new URLSearchParams({q, sheet, risk, overdue});
    const rows = await api("/api/assets?" + params.toString());
    state.assets = rows;
    const el = document.getElementById("assetsTableWrap");
    el.innerHTML = tableHtml(rows, false);
    bindRowClicks(el);
  } catch (err) {
    showToast("Failed to load assets: " + err.message);
  }
}

// ---------------------------------------------------------------- critical replacements
document.getElementById("criticalSearch")?.addEventListener("input", loadCriticalAssets);
document.getElementById("criticalDoneFilter")?.addEventListener("change", loadCriticalAssets);

async function loadCriticalAssets() {
  try {
    const q = (document.getElementById("criticalSearch")?.value || "").toLowerCase();
    const doneFilter = document.getElementById("criticalDoneFilter")?.value || "";
    let rows = await api("/api/critical_assets");
    
    if (q) {
      rows = rows.filter(r => (r.item_description || "").toLowerCase().includes(q) ||
                              (r.replacement_scope || "").toLowerCase().includes(q) ||
                              (r.pack_no || "").toLowerCase().includes(q) ||
                              (r.remarks || "").toLowerCase().includes(q));
    }
    if (doneFilter) {
      rows = rows.filter(r => (r.replacement_done || "").toLowerCase() === doneFilter.toLowerCase());
    }

    const el = document.getElementById("criticalTableWrap");
    if (!rows.length) {
      el.innerHTML = `<table class="assets"><tbody><tr class="empty-row"><td>No matching critical replacement records.</td></tr></tbody></table>`;
      return;
    }

    el.innerHTML = `<table class="assets">
      <thead>
        <tr><th>SN</th><th>Pack #</th><th>Item Description</th><th>Category Scope</th><th>Replacement Scope</th><th>Done?</th><th>Remarks</th></tr>
      </thead>
      <tbody>${rows.map(c => `<tr>
        <td>${esc(c.sn || "—")}</td>
        <td><b>${esc(c.pack_no || "—")}</b></td>
        <td><b>${esc(c.item_description || "—")}</b></td>
        <td><span class="tag-mono">${esc(c.category_section || "")}</span></td>
        <td><div style="max-width:280px; font-size:12px;">${esc(c.replacement_scope || "")}</div></td>
        <td>
          <button class="btn ${c.replacement_done === "Yes" ? "ghost" : "primary"} small" onclick="toggleCriticalDone(event, ${c.id}, '${c.replacement_done === "Yes" ? "No" : "Yes"}')">
            ${c.replacement_done === "Yes" ? "Completed (Yes)" : "Pending (No)"}
          </button>
        </td>
        <td><div style="font-size:11.5px; color:var(--text-muted);">${esc(c.remarks || "")} ${c.plant_remarks ? " | " + esc(c.plant_remarks) : ""}</div></td>
      </tr>`).join("")}</tbody>
    </table>`;
  } catch (err) {
    showToast("Failed to load critical assets: " + err.message);
  }
}

window.toggleCriticalDone = async function(event, id, doneStatus) {
  if (event) event.stopPropagation();
  try {
    await api("/api/critical_assets/" + id, {
      method: "PUT",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({replacement_done: doneStatus}),
    });
    loadCriticalAssets();
    showToast(`Replacement status set to ${doneStatus}`, false);
  } catch (err) {
    showToast("Failed to update status: " + err.message);
  }
};

// ---------------------------------------------------------------- temporary repairs
document.getElementById("tempRepairSearch")?.addEventListener("input", loadTempRepairs);
document.getElementById("tempRepairStatusFilter")?.addEventListener("change", loadTempRepairs);

async function loadTempRepairs() {
  try {
    const q = (document.getElementById("tempRepairSearch")?.value || "").toLowerCase();
    const statusFilter = document.getElementById("tempRepairStatusFilter")?.value || "";
    let rows = await api("/api/temp_repairs");

    if (q) {
      rows = rows.filter(r => (r.asset_name || "").toLowerCase().includes(q) ||
                              (r.repaired_section || "").toLowerCase().includes(q) ||
                              (r.area || "").toLowerCase().includes(q) ||
                              (r.repaired_by || "").toLowerCase().includes(q) ||
                              (r.report_ref || "").toLowerCase().includes(q));
    }
    if (statusFilter) {
      rows = rows.filter(r => (r.expiration_status || "").toLowerCase() === statusFilter.toLowerCase());
    }

    const el = document.getElementById("tempRepairsTableWrap");
    if (!rows.length) {
      el.innerHTML = `<table class="assets"><tbody><tr class="empty-row"><td>No matching temporary repairs.</td></tr></tbody></table>`;
      return;
    }

    el.innerHTML = `<table class="assets">
      <thead>
        <tr><th>Facility / Area</th><th>Asset & Section</th><th>Repaired By (Type)</th><th>Orig Date</th><th>Life (Yrs)</th><th>Expire Date</th><th>Hardness (HB)</th><th>Status</th></tr>
      </thead>
      <tbody>${rows.map(r => {
        const isExp = (r.expiration_status || "").toLowerCase() === "expired" || (r.expiration_date && r.expiration_date < new Date().toISOString().split("T")[0]);
        const pillCls = (r.expiration_status || "").toLowerCase() === "replaced" ? "ok" : (isExp ? "overdue" : "soon");
        return `<tr>
          <td><div><b>${esc(r.area || "")}</b></div><div class="tag-mono">${esc(r.facility_type || "")}</div></td>
          <td><div><b>${esc(r.asset_name || "")}</b></div><div style="font-size:12px; color:var(--accent);">${esc(r.repaired_section || "")}</div></td>
          <td>${esc(r.repaired_by || "—")}</td>
          <td>${esc(r.original_repair_date || "—")}</td>
          <td>${esc(r.repair_life_years || "—")}</td>
          <td><b>${esc(r.expiration_date || "—")}</b></td>
          <td>${esc(r.hardness_hb || "—")}</td>
          <td><span class="status-pill ${pillCls}">${esc(r.expiration_status || "Active")}</span></td>
        </tr>`;
      }).join("")}</tbody>
    </table>`;
  } catch (err) {
    showToast("Failed to load temporary repairs: " + err.message);
  }
}

// ---------------------------------------------------------------- yearly plan
document.getElementById("btnLoadYear").addEventListener("click", async () => {
  const year = document.getElementById("yearInput").value;
  const rows = await api("/api/yearly_plan?year=" + encodeURIComponent(year));
  document.getElementById("yearlyMeta").textContent =
    `${rows.length} asset(s) with a due date falling in ${year}.`;
  const el = document.getElementById("yearlyTableWrap");
  el.innerHTML = tableHtml(rows, false);
  bindRowClicks(el);
});
document.getElementById("btnPrintYear").addEventListener("click", () => window.print());

// ---------------------------------------------------------------- export & re-import
document.getElementById("btnExport").addEventListener("click", () => {
  const sheet = document.getElementById("sheetFilter").value;
  const params = new URLSearchParams(sheet ? {sheet} : {});
  window.location.href = "/api/export.csv?" + params.toString();
});

const picker = document.getElementById("excelFilePicker");

document.getElementById("btnReimport").addEventListener("click", () => {
  picker.value = "";
  picker.click();
});

picker.addEventListener("change", async () => {
  const file = picker.files[0];
  if (!file) return;

  const msg = `Are you sure you want to import "${file.name}"?\n\nThis will reset the database with the assets from this Excel file.`;
  if (!confirm(msg)) return;

  const btn = document.getElementById("btnReimport");
  const origText = btn.textContent;
  btn.disabled = true;
  btn.textContent = "Importing…";

  try {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const base64Data = e.target.result.split(",")[1];
        const res = await api("/api/reimport_file", {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify({filename: file.name, filedata: base64Data}),
        });
        alert(res.message || "Re-import successful!");
        refreshCurrentView();
      } catch (err) {
        alert("Re-import failed: " + err.message);
      } finally {
        btn.disabled = false;
        btn.textContent = origText;
      }
    };
    reader.onerror = () => {
      alert("Failed to read file.");
      btn.disabled = false;
      btn.textContent = origText;
    };
    reader.readAsDataURL(file);
  } catch (err) {
    alert("Error preparing file upload: " + err.message);
    btn.disabled = false;
    btn.textContent = origText;
  }
});

// ---------------------------------------------------------------- drawer: view/edit asset
const overlay = document.getElementById("drawerOverlay");
document.getElementById("drawerClose").addEventListener("click", closeDrawer);
overlay.addEventListener("click", e => { if (e.target === overlay) closeDrawer(); });

function closeDrawer() {
  overlay.classList.remove("open");
  state.editingAsset = null;
}

document.addEventListener("keydown", e => {
  if (e.key === "Escape") closeDrawer();
});

const CORE_FIELDS = [
  ["name", "Asset name"], ["field", "Pack # / Field Section"], ["location", "Location / Section"],
  ["tag", "Tag #"], ["asset_number", "Asset #"], ["plant", "Plant / Facility"],
  ["unit_name", "Unit name"], ["source_sheet", "Source sheet"], ["in_service", "In service?"],
  ["insulation", "Insulation"],
];

const ENVELOPE_FIELDS = [
  ["fluid_service", "Fluid Service"], ["material_spec", "Material Spec (e.g. CS A106)"],
  ["design_pressure", "Design Press (psi/bar)"], ["operating_pressure", "Operating Press"],
  ["design_temp", "Design Temp (°C)"], ["operating_temp", "Operating Temp (°C)"],
  ["nominal_thickness", "Nominal Thickness (mm)"], ["t_min", "Min Allowable t_min (mm)"],
];

const RISK_FIELDS = [
  ["risk_category", "Risk Level (HIGH / MEDIUM / LOW)"],
  ["damage_mechanisms", "Damage Mechanisms (e.g. CUI, Thinning)"],
  ["corrosion_rate", "Worst Corrosion Rate (mm/yr)"],
  ["remaining_life", "Remaining Life (years)"],
];

const DATE_FIELDS = [
  ["date_osi_last", "Last OSI date"], ["date_osi_next", "Next OSI due"],
  ["date_internal_last", "Last internal date"], ["date_internal_next", "Next internal due"],
];

function fieldInput(key, label, value, type) {
  const val = esc(value ?? "");
  const inputType = type || "text";
  return `<div class="field">
    <label>${label}</label>
    <input data-field="${key}" type="${inputType}" value="${val}">
  </div>`;
}

async function openAssetDrawer(id) {
  const a = await api("/api/assets/" + id);
  state.editingAsset = a;
  document.getElementById("drawerTitle").innerHTML = `${esc(a.name || "Asset")} ${riskBadge(a)}`;

  const extraRows = Object.entries(a.extra || {}).filter(([k]) => !k.endsWith("(raw)"));

  document.getElementById("drawerBody").innerHTML = `
    <div class="drawer-tabs">
      <button class="drawer-tab active" onclick="switchDrawerTab('specs')">Specs & Envelope</button>
      <button class="drawer-tab" onclick="switchDrawerTab('risk')">Risk & Condition</button>
      <button class="drawer-tab" onclick="switchDrawerTab('logs')">NDT & Inspection Logs</button>
    </div>

    <!-- TAB 1: SPECS -->
    <div id="dtab-specs" class="dtab-content">
      <div class="section-title">Identification & Location</div>
      <div class="field-grid">
        ${CORE_FIELDS.map(([k,l]) => fieldInput(k,l,a[k])).join("")}
      </div>

      <div class="section-title">Design & Operating Envelope</div>
      <div class="field-grid">
        ${ENVELOPE_FIELDS.map(([k,l]) => fieldInput(k,l,a[k])).join("")}
      </div>
    </div>

    <!-- TAB 2: RISK & CONDITION -->
    <div id="dtab-risk" class="dtab-content" style="display:none;">
      <div class="section-title">Risk & Damage Mechanisms</div>
      <div class="field-grid">
        ${RISK_FIELDS.map(([k,l]) => fieldInput(k,l,a[k])).join("")}
      </div>

      <div class="section-title">Scheduled Inspection Dates</div>
      <div class="field-grid">
        ${DATE_FIELDS.map(([k,l]) => fieldInput(k,l,a[k],"text")).join("")}
      </div>

      <div class="field-grid wide">
        <div class="field"><label>Remarks & Integrity Recommendations</label>
          <textarea data-field="remarks">${esc(a.remarks)}</textarea></div>
      </div>

      ${extraRows.length ? `
        <div class="section-title">Original Sheet Metadata</div>
        <div class="field-grid wide">
          ${extraRows.map(([k,v]) => `<div class="field"><label>${esc(k)}</label>
            <input value="${esc(v)}" readonly></div>`).join("")}
        </div>` : ""}
    </div>

    <!-- TAB 3: LOGS & NDT -->
    <div id="dtab-logs" class="dtab-content" style="display:none;">
      <div class="section-title">Log New NDT / Inspection</div>
      <div class="field-grid">
        <div class="field"><label>Inspection Date</label><input id="logDate" type="date"></div>
        <div class="field"><label>Inspection Type</label>
          <select id="logType">
            <option>OSI</option><option>Internal</option><option>OSI-ADV</option>
            <option>NDT UT Thickness</option><option>Visual / CUI</option><option>Other</option>
          </select>
        </div>
      </div>
      <div class="field-grid">
        <div class="field"><label>Inspector / Cert #</label><input id="logInspector" placeholder="e.g. J. Doe (API 510 #1234)"></div>
        <div class="field"><label>NDT Method</label><input id="logMethod" placeholder="e.g. UT, RT, MPT, Visual"></div>
      </div>
      <div class="field-grid">
        <div class="field"><label>Actual Min Thickness t_actual (mm)</label><input id="logTActual" placeholder="e.g. 6.2 mm"></div>
        <div class="field"><label>Corrective Action Required?</label><input id="logAction" placeholder="e.g. Yes / None"></div>
      </div>
      <div class="field-grid wide">
        <div class="field"><label>Findings & Observations</label><textarea id="logFindings"></textarea></div>
      </div>
      <div class="field-grid">
        <div class="field"><label>Next Scheduled Due Date</label><input id="logNextDue" type="date"></div>
      </div>
      <div class="drawer-actions">
        <button id="btnAddLog" class="btn primary">Record Inspection</button>
      </div>

      ${a.log && a.log.length ? `
        <div class="section-title">Historical Inspection & NDT Log</div>
        ${a.log.map(l => `<div class="log-entry" style="position:relative;">
          <div class="meta">${esc(l.insp_date || "")} · ${esc(l.insp_type || "")}${l.insp_method ? " · " + esc(l.insp_method) : ""}${l.next_due_date ? " · next due " + esc(l.next_due_date) : ""}</div>
          ${l.inspector_name ? `<div style="font-weight:600; font-size:11.5px; color:var(--accent);">Inspector: ${esc(l.inspector_name)}</div>` : ""}
          ${l.t_actual ? `<div style="font-size:11.5px; color:var(--text-muted);">Measured Wall: ${esc(l.t_actual)}</div>` : ""}
          <div style="margin-top:4px;">${esc(l.findings || "")}</div>
          <button class="btn danger small" style="position:absolute; top:8px; right:8px; padding:2px 6px; font-size:11px;" onclick="deleteLog(${a.id}, ${l.id})">Delete</button>
        </div>`).join("")}
      ` : ""}
    </div>

    <div class="drawer-actions" style="margin-top:20px; border-top:1px solid var(--border); padding-top:12px;">
      <button id="btnSaveAsset" class="btn primary">Save Changes</button>
      <button id="btnArchiveAsset" class="btn danger">${a.archived ? "Restore" : "Archive"}</button>
    </div>
  `;

  document.getElementById("btnSaveAsset").addEventListener("click", () => saveAsset(id));
  document.getElementById("btnArchiveAsset").addEventListener("click", () => toggleArchive(id, !a.archived));
  document.getElementById("btnAddLog").addEventListener("click", () => addLog(id));

  overlay.classList.add("open");
}

window.switchDrawerTab = function(tabName) {
  document.querySelectorAll(".drawer-tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".dtab-content").forEach(c => c.style.display = "none");
  
  event.target.classList.add("active");
  const el = document.getElementById("dtab-" + tabName);
  if (el) el.style.display = "block";
};

async function saveAsset(id) {
  const inputs = document.querySelectorAll("#drawerBody [data-field]");
  const payload = {};
  inputs.forEach(el => { payload[el.dataset.field] = el.value || null; });
  try {
    await api("/api/assets/" + id, {
      method: "PUT",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload),
    });
    closeDrawer();
    refreshCurrentView();
  } catch (err) {
    showToast("Save failed: " + err.message);
  }
}

async function toggleArchive(id, archived) {
  try {
    await api("/api/assets/" + id, {
      method: "PUT",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({archived: archived ? 1 : 0}),
    });
    closeDrawer();
    refreshCurrentView();
  } catch (err) {
    showToast("Archive toggle failed: " + err.message);
  }
}

async function addLog(id) {
  const payload = {
    insp_date: document.getElementById("logDate").value || null,
    insp_type: document.getElementById("logType").value,
    inspector_name: document.getElementById("logInspector").value || null,
    insp_method: document.getElementById("logMethod").value || null,
    t_actual: document.getElementById("logTActual").value || null,
    action_required: document.getElementById("logAction").value || null,
    findings: document.getElementById("logFindings").value || null,
    next_due_date: document.getElementById("logNextDue").value || null,
  };
  try {
    await api("/api/assets/" + id + "/log", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload),
    });
    openAssetDrawer(id); // refresh drawer with new log entry
    refreshCurrentView();
  } catch (err) {
    showToast("Failed to add log: " + err.message);
  }
}

window.deleteLog = async function(assetId, logId) {
  if (!confirm("Are you sure you want to delete this inspection log entry?")) return;
  try {
    await api("/api/logs/" + logId, { method: "DELETE" });
    openAssetDrawer(assetId);
    refreshCurrentView();
    showToast("Inspection record removed", false);
  } catch (err) {
    showToast("Failed to delete log: " + err.message);
  }
};

function refreshCurrentView() {
  if (state.view === "dashboard") loadDashboard();
  if (state.view === "assets") loadAssets();
}

// ---------------------------------------------------------------- new asset
document.getElementById("btnNewAsset").addEventListener("click", () => {
  state.editingAsset = null;
  document.getElementById("drawerTitle").textContent = "New Asset";
  document.getElementById("drawerBody").innerHTML = `
    <div class="section-title">Identification & Location</div>
    <div class="field-grid">
      ${CORE_FIELDS.map(([k,l]) => fieldInput(k,l,k==="source_sheet"?"Manual Entry":"")).join("")}
    </div>
    <div class="section-title">Design & Operating Envelope</div>
    <div class="field-grid">
      ${ENVELOPE_FIELDS.map(([k,l]) => fieldInput(k,l,"")).join("")}
    </div>
    <div class="section-title">Risk Profile</div>
    <div class="field-grid">
      ${RISK_FIELDS.map(([k,l]) => fieldInput(k,l,"")).join("")}
    </div>
    <div class="section-title">Inspection dates</div>
    <div class="field-grid">
      ${DATE_FIELDS.map(([k,l]) => fieldInput(k,l,"","date")).join("")}
    </div>
    <div class="field-grid wide">
      <div class="field"><label>Remarks & Integrity Recommendations</label><textarea data-field="remarks"></textarea></div>
    </div>
    <div class="drawer-actions">
      <button id="btnCreateAsset" class="btn primary">Create asset</button>
    </div>
  `;
  document.getElementById("btnCreateAsset").addEventListener("click", createAsset);
  overlay.classList.add("open");
});

async function createAsset() {
  const inputs = document.querySelectorAll("#drawerBody [data-field]");
  const payload = {};
  inputs.forEach(el => { payload[el.dataset.field] = el.value || null; });
  const a = await api("/api/assets", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify(payload),
  });
  closeDrawer();
  refreshCurrentView();
  openAssetDrawer(a.id);
}

// ---------------------------------------------------------------- init
loadDashboard();