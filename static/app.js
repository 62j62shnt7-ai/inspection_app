/**
 * Master Inspection Plan & Asset Integrity Management Suite
 */

const state = {
  view: "dashboard",
  assets: [],
  filteredAssets: [],
  selectedAssetIds: new Set(),
  page: 1,
  pageSize: 50,
  groupBy: "",
  sheets: [],
  activeRbiCell: "",
  activeAgingBucket: "",
  editingAsset: null,
  activeDrawerTab: "specs",
};

// ---------------------------------------------------------------- Toast Notifications
function showToast(msg, isError = false) {
  let container = document.getElementById("toastContainer");
  if (!container) {
    container = document.createElement("div");
    container.id = "toastContainer";
    document.body.appendChild(container);
  }
  const el = document.createElement("div");
  el.className = "toast" + (isError ? " error" : "");
  el.innerHTML = `<span>${isError ? "⚠️" : "✅"}</span> <span>${esc(msg)}</span>`;
  container.appendChild(el);
  requestAnimationFrame(() => el.classList.add("show"));
  setTimeout(() => {
    el.classList.remove("show");
    setTimeout(() => el.remove(), 300);
  }, 4500);
}

// ---------------------------------------------------------------- Helpers
function esc(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/[&<>"']/g, c => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  })[c]);
}

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
  if (asset.due_date_estimated && asset.overdue) {
    return `<span class="status-pill deferred" title="Due date is an import placeholder (original cell had no parseable date) — excluded from aging stats">⚠️ ESTIMATED DATE</span>`;
  }
  if (asset.is_deferred) {
    return `<span class="status-pill deferred" title="Approved MOC Deferral until ${asset.deferral_expiry || ''}">🛡️ DEFERRED (${fmtDate(asset.deferral_expiry)})</span>`;
  }
  if (asset.is_next_sd) {
    return `<span class="status-pill warning" title="Scheduled for Next Shutdown">⚠️ NEXT SD</span>`;
  }
  if (asset.is_waiting_eval) {
    return `<span class="status-pill info" title="Waiting for Evaluation">ℹ️ W/EVAL</span>`;
  }
  if (asset.status_osi_next === "OOS" || asset.status_internal_next === "OOS") {
    return `<span class="status-pill none" title="Out of Service">OUT OF SERVICE</span>`;
  }

  // Check overdue flags
  if (asset.is_overdue) {
    let overdueLabel = "OVERDUE";
    if (asset.overdue_type === "Both") {
      overdueLabel = "OVERDUE (Both)";
    } else if (asset.overdue_type === "Internal") {
      overdueLabel = "OVERDUE (Internal SD)";
    } else if (asset.overdue_type === "OSI") {
      overdueLabel = "OVERDUE (OSI)";
    }
    const days = daysFromToday(asset.next_due);
    const dayStr = days !== null ? ` ${Math.abs(days)}d` : "";
    return `<span class="status-pill overdue" title="Inspection Overdue">${overdueLabel}${dayStr}</span>`;
  }

  const nd = asset.next_due;
  if (!nd) {
    if (asset.status_osi_next || asset.status_internal_next) {
      return `<span class="status-pill none">${esc(asset.status_osi_next || asset.status_internal_next)}</span>`;
    }
    return `<span class="status-pill none">No due date</span>`;
  }
  const days = daysFromToday(nd);
  if (days !== null && days <= 30) return `<span class="status-pill soon">DUE ${days}d</span>`;
  return `<span class="status-pill ok">${fmtDate(nd)}</span>`;
}

function riskBadge(asset) {
  const r = (asset.risk_category || "LOW").toUpperCase();
  const cls = r === "HIGH" ? "high" : (r === "MEDIUM" ? "medium" : "low");
  const cuiHtml = asset.is_cui ? `<span class="badge-cui" title="Corrosion Under Insulation Susceptible">CUI</span>` : "";
  const rbiCell = asset.rbi_matrix_cell ? `<span class="badge-rbi-cell" title="RBI 5x5 Matrix Cell">${asset.rbi_matrix_cell}</span>` : "";
  return `<span class="badge-risk ${cls}">${r}</span>${rbiCell}${cuiHtml}`;
}

async function api(path, opts) {
  const res = await fetch(path, opts);
  if (!res.ok) {
    const err = await res.json().catch(() => ({error: res.statusText}));
    throw new Error(err.error || "Request failed");
  }
  return res.json();
}

// ---------------------------------------------------------------- Navigation Tabs
document.querySelectorAll(".tab").forEach(btn => {
  btn.addEventListener("click", () => switchView(btn.dataset.view));
});

function switchView(view) {
  state.view = view;
  document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b.dataset.view === view));
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  
  const targetView = document.getElementById("view-" + view);
  if (targetView) targetView.classList.add("active");

  if (view === "dashboard") loadDashboard();
  if (view === "assets") loadAssets();
  if (view === "critical") loadCriticalAssets();
  if (view === "temprepairs") loadTempRepairs();
  if (view === "yearly") {
    if (!document.getElementById("yearInput").value) {
      document.getElementById("yearInput").value = new Date().getFullYear();
    }
    const ysel = document.getElementById("yearlySheetFilter");
    if (ysel && (!ysel.options || ysel.options.length <= 1) && state.sheets && state.sheets.length) {
      const currentY = ysel.value;
      ysel.innerHTML = `<option value="">All Sheets (Global)</option>` + 
        state.sheets.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
      ysel.value = currentY;
    }
    loadYearlyPlan();
  }
  if (view === "reconciliation") {
    loadReconciliation();
  }
}

// ---------------------------------------------------------------- Dashboard & Asset Table
function assetRowHtml(a, compact) {
  const tag = a.tag || a.sn || "";
  const loc = [a.plant, a.location].filter(Boolean).join(" · ") || "—";
  const isSelected = state.selectedAssetIds.has(a.id);
  const checkTd = compact ? "" : `<td style="width:36px; text-align:center;" onclick="event.stopPropagation();">
    <input type="checkbox" class="row-checkbox" data-id="${a.id}" ${isSelected ? "checked" : ""}>
  </td>`;

  if (compact) {
    return `<tr data-id="${a.id}">
      <td>
        <div style="font-weight:600; color:#fff;">${esc(a.name || "(unnamed)")} ${riskBadge(a)}</div>
        <div class="tag-mono">${esc(tag)}</div>
      </td>
      <td><span class="tag-mono">${esc(a.source_sheet || "")}</span></td>
      <td>${statusPill(a)}</td>
    </tr>`;
  }

  const press = a.operating_pressure ? `${esc(a.operating_pressure)} psi` : (a.design_pressure ? `${esc(a.design_pressure)} psi (des)` : "—");

  // Format OSI timeline
  const osiLast = a.date_osi_last || a.date_osi_last_raw;
  const osiNext = a.date_osi_next || a.date_osi_next_raw || a.status_osi_next;
  const osiHtml = (osiLast || osiNext) ? `
    <div style="font-size:11.5px; line-height:1.4;">
      <span class="badge-scope osi">OSI</span>
      <span style="color:var(--text-muted);">Last:</span> <b style="color:#fff;">${fmtDate(osiLast)}</b>
      <span style="color:var(--text-muted); margin-left:4px;">Next:</span> <b style="color:${a.osi_overdue ? 'var(--danger)' : 'var(--accent)'};">${fmtDate(osiNext)}</b>
    </div>` : "";

  // Format Internal timeline
  const intLast = a.date_internal_last || a.date_internal_last_raw;
  const intNext = a.date_internal_next || a.date_internal_next_raw || a.status_internal_next;
  const intHtml = (intLast || intNext) ? `
    <div style="font-size:11.5px; line-height:1.4; margin-top:2px;">
      <span class="badge-scope internal">Internal</span>
      <span style="color:var(--text-muted);">Last:</span> <b style="color:#fff;">${fmtDate(intLast)}</b>
      <span style="color:var(--text-muted); margin-left:4px;">Next:</span> <b style="color:${a.internal_overdue ? 'var(--danger)' : 'var(--accent)'};">${fmtDate(intNext)}</b>
    </div>` : "";

  const planScopeHtml = a.plan_insp_type ? `
    <div style="margin-bottom:3px;">
      <span class="badge-scope ${a.plan_insp_type.toLowerCase().includes('internal') ? 'internal' : (a.plan_insp_type.toLowerCase().includes('osi') ? 'osi' : 'both')}">
        🗓️ Planned: ${esc(a.plan_insp_type)} (${fmtDate(a.plan_date)})
      </span>
      ${a.plan_is_carry_over ? `<span class="badge-scope" style="background:#7f1d1d; color:#fecaca; margin-left:4px;">⚠ OVERDUE CARRY-OVER</span>` : ""}
      ${a.plan_is_deferred ? `<span class="badge-scope" style="background:#78350f; color:#fde68a; margin-left:4px;">DEFERRED</span>` : ""}
      ${a.plan_is_estimated && !a.plan_is_carry_over ? `<span class="badge-scope" style="background:#374151; color:#d1d5db; margin-left:4px;">ESTIMATED DATE</span>` : ""}
    </div>` : "";

  const timelineContent = (osiHtml || intHtml) ? (planScopeHtml + osiHtml + intHtml) : `<span style="color:var(--text-muted); font-size:12px;">No schedule recorded</span>`;

  return `<tr data-id="${a.id}">
    ${checkTd}
    <td>
      <div style="font-weight:600; color:#fff;">${esc(a.name || "(unnamed)")} ${riskBadge(a)}</div>
      <div class="tag-mono">${esc(tag)}</div>
    </td>
    <td><span class="tag-mono" style="font-weight:600; color:var(--accent);">${esc(a.source_sheet || "")}</span></td>
    <td>${statusPill(a)}</td>
    <td><span class="tag-mono" style="font-weight:600; color:#fff;">${esc(a.field || "—")}</span></td>
    <td>${timelineContent}</td>
    <td>
      <div style="font-size:12px; color:#fff;">${press}</div>
      ${a.nominal_thickness ? `<div style="font-size:11px; color:var(--text-faint);">Thk: ${esc(a.nominal_thickness)} mm</div>` : ""}
    </td>
    <td><div style="max-width:240px; font-size:12px; color:var(--text-muted); overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${esc(a.remarks || loc)}">${esc(a.remarks || loc)}</div></td>
  </tr>`;
}

function tableHtml(rows, compact) {
  if (!rows || !rows.length) {
    return `<table class="assets"><tbody><tr class="empty-row"><td>No records found.</td></tr></tbody></table>`;
  }
  const checkTh = compact ? "" : `<th style="width:36px; text-align:center;"><input type="checkbox" id="selectAllCheckbox"></th>`;
  const extraHeads = compact ? "" : `<th>Pack # / Field</th><th>Inspection Timeline</th><th>Pressure / Thk</th><th>Remarks & Location</th>`;
  return `<table class="assets">
    <thead><tr>${checkTh}<th>Asset & Risk</th><th>Source Sheet</th><th>Status</th>${extraHeads}</tr></thead>
    <tbody>${rows.map(a => assetRowHtml(a, compact)).join("")}</tbody>
  </table>`;
}

function bindRowClicks(container) {
  if (!container) return;
  container.querySelectorAll("tr[data-id]").forEach(tr => {
    tr.addEventListener("click", () => openAssetDrawer(parseInt(tr.dataset.id)));
  });

  // Checkbox interactions
  container.querySelectorAll(".row-checkbox").forEach(cb => {
    cb.addEventListener("change", (e) => {
      e.stopPropagation();
      const id = parseInt(cb.dataset.id);
      if (cb.checked) state.selectedAssetIds.add(id);
      else state.selectedAssetIds.delete(id);
      updateBatchActionBar();
    });
  });

  const selectAll = container.querySelector("#selectAllCheckbox");
  if (selectAll) {
    selectAll.addEventListener("change", (e) => {
      e.stopPropagation();
      const rows = state.filteredAssets;
      if (selectAll.checked) {
        rows.forEach(r => state.selectedAssetIds.add(r.id));
      } else {
        rows.forEach(r => state.selectedAssetIds.delete(r.id));
      }
      container.querySelectorAll(".row-checkbox").forEach(cb => cb.checked = selectAll.checked);
      updateBatchActionBar();
    });
  }
}

// ---------------------------------------------------------------- Audit & Backups
async function loadAuditTrail() {
  const el = document.getElementById("auditList");
  if (!el) return;
  try {
    const rows = await api("/api/audit");
    el.innerHTML = rows.length ? `<table class="assets"><tbody>
      ${rows.slice(0, 20).map(r => `<tr>
        <td style="white-space:nowrap; color:var(--text-faint);">${esc(r.ts)}</td>
        <td style="white-space:nowrap;"><b>${esc(r.action)}</b></td>
        <td>${esc(r.entity_type || "")}${r.entity_id ? " #" + esc(r.entity_id) : ""}</td>
        <td style="color:var(--text-muted); font-size:11.5px;">${esc(r.details || "")}</td>
      </tr>`).join("")}</tbody></table>`
      : `<div style="color:var(--text-faint); font-size:12.5px; padding:10px;">No activity recorded yet. Every log, deferral, sync and restore will appear here.</div>`;
  } catch { el.innerHTML = "<div style='color:var(--text-faint); padding:10px;'>Audit log unavailable.</div>"; }
}

async function loadBackups() {
  const el = document.getElementById("backupList");
  if (!el) return;
  try {
    const rows = await api("/api/backups");
    el.innerHTML = rows.length ? `<table class="assets"><tbody>
      ${rows.slice(0, 12).map(b => `<tr>
        <td style="white-space:nowrap;">💾 ${esc(b.filename)}</td>
        <td style="white-space:nowrap; color:var(--text-faint);">${esc(b.modified)} · ${b.size_mb} MB</td>
        <td style="text-align:right;"><button class="btn ghost small" onclick="restoreBackup('${esc(b.filename)}')">Restore</button></td>
      </tr>`).join("")}</tbody></table>`
      : `<div style="color:var(--text-faint); font-size:12.5px; padding:10px;">No backups yet. One is taken automatically before sync-to-master, reimport or database clear.</div>`;
  } catch { el.innerHTML = "<div style='color:var(--text-faint); padding:10px;'>Backup list unavailable.</div>"; }
}

window.restoreBackup = async function(filename) {
  if (!confirm(`Restore database from "${filename}"?\n\nYour current data will be backed up first, then replaced.`)) return;
  try {
    const res = await api("/api/backups/restore", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({filename})
    });
    showToast(res.message || "Restored.");
    loadDashboard();
  } catch (err) {
    showToast("Restore failed: " + err.message, true);
  }
};

// ---------------------------------------------------------------- 5x5 RBI Matrix Heatmap
const MATRIX_CELL_TIERS = {
  "5E": "rbi-high", "5D": "rbi-high", "4E": "rbi-high",
  "5C": "rbi-medhigh", "4D": "rbi-medhigh", "3E": "rbi-medhigh", "5B": "rbi-medhigh",
  "5A": "rbi-med", "4C": "rbi-med", "3D": "rbi-med", "2E": "rbi-med", "4B": "rbi-med", "3C": "rbi-med", "2D": "rbi-med", "1E": "rbi-med",
  "4A": "rbi-low", "3B": "rbi-low", "3A": "rbi-low", "2C": "rbi-low", "2B": "rbi-low", "2A": "rbi-low", "1D": "rbi-low", "1C": "rbi-low", "1B": "rbi-low", "1A": "rbi-low"
};

function render5x5Matrix(matrixCounts) {
  const container = document.getElementById("rbiMatrixContainer");
  if (!container) return;

  const pofLabels = ["A (Very Low)", "B (Low)", "C (Moderate)", "D (High)", "E (Very High)"];
  const pofLetters = ["A", "B", "C", "D", "E"];
  const cofRows = [
    {val: 5, lbl: "5 (Catastrophic)"},
    {val: 4, lbl: "4 (Major)"},
    {val: 3, lbl: "3 (Moderate)"},
    {val: 2, lbl: "2 (Minor)"},
    {val: 1, lbl: "1 (Negligible)"}
  ];

  let html = `<table class="matrix-table">
    <thead>
      <tr>
        <th class="matrix-axis-label" style="text-align:right;">COF \\ POF</th>
        ${pofLabels.map((l, idx) => `<th class="matrix-axis-label">${pofLetters[idx]}<br><span style="font-size:9.5px; opacity:0.7;">${l.split(" ")[1]}</span></th>`).join("")}
      </tr>
    </thead>
    <tbody>`;

  cofRows.forEach(crow => {
    html += `<tr><td class="matrix-axis-label" style="text-align:right; padding-right:8px;">${crow.val}<br><span style="font-size:9.5px; opacity:0.7;">${crow.lbl.split(" ")[1]}</span></td>`;
    pofLetters.forEach(plet => {
      const cellKey = `${crow.val}${plet}`;
      const count = matrixCounts[cellKey] || 0;
      const tierClass = MATRIX_CELL_TIERS[cellKey] || "rbi-low";
      const isActive = state.activeRbiCell === cellKey ? "active-filter" : "";
      
      html += `<td class="matrix-cell ${tierClass} ${isActive}" onclick="filterByRbiCell('${cellKey}')" title="Filter by Cell ${cellKey} (${count} assets)">
        <div class="matrix-cell-code">${cellKey}</div>
        <div class="matrix-cell-count">${count}</div>
      </td>`;
    });
    html += `</tr>`;
  });

  html += `</tbody></table>`;
  container.innerHTML = html;
}

window.filterByRbiCell = function(cellKey) {
  state.activeRbiCell = cellKey;
  const btnReset = document.getElementById("btnResetMatrixFilter");
  if (btnReset) btnReset.style.display = "block";
  switchView("assets");
};

window.clearMatrixFilter = function() {
  state.activeRbiCell = "";
  const btnReset = document.getElementById("btnResetMatrixFilter");
  if (btnReset) btnReset.style.display = "none";
  loadDashboard();
};

// ---------------------------------------------------------------- Dashboard Loader
async function loadDashboard() {
  try {
    const d = await api("/api/dashboard");
    
    // 1. Executive Compliance Banner
    const kpiRate = document.getElementById("kpiComplianceRate");
    if (kpiRate) kpiRate.textContent = `${d.compliance_rate}%`;
    const kpiRateLbl = document.getElementById("kpiComplianceLbl");
    if (kpiRateLbl) {
      const unrated = d.no_due_date_count || 0;
      kpiRateLbl.textContent = unrated
        ? `Inspection Compliance (${unrated} asset${unrated === 1 ? "" : "s"} unrated — no due date)`
        : "Inspection Compliance (all assets scheduled)";
    }
    const kpiUnmanaged = document.getElementById("kpiUnmanagedOverdue");
    if (kpiUnmanaged) kpiUnmanaged.textContent = d.aging_counts?.unmanaged_overdue ?? d.overdue_count;
    const kpiApproved = document.getElementById("kpiApprovedDeferrals");
    if (kpiApproved) kpiApproved.textContent = d.aging_counts?.approved_deferrals ?? 0;
    const kpiCui = document.getElementById("kpiCuiCount");
    if (kpiCui) kpiCui.textContent = d.cui_count;

    // 2. Stat Cards
    const statRow = document.getElementById("statRow");
    if (statRow) {
      statRow.innerHTML = `
        <div class="stat-card">
          <div class="n">${d.total_assets}</div>
          <div class="label">Total Tracked Assets</div>
        </div>
        <div class="stat-card overdue" onclick="filterByOverdueType('')" style="cursor:pointer;" title="Click to view all overdue">
          <div class="n">${d.overdue_count}</div>
          <div class="label">Total Overdue</div>
        </div>
        <div class="stat-card overdue" onclick="filterByOverdueType('internal')" style="cursor:pointer; border-color:#e05252;" title="Click to view Internal Overdue (requires Shutdown)">
          <div class="n">${d.overdue_internal_count || 0}</div>
          <div class="label">Internal Overdue (SD)</div>
        </div>
        <div class="stat-card overdue" onclick="filterByOverdueType('osi')" style="cursor:pointer; border-color:#ff9800;" title="Click to view OSI Overdue (On-Stream)">
          <div class="n">${d.overdue_osi_count || 0}</div>
          <div class="label">OSI Overdue (On-Stream)</div>
        </div>
        <div class="stat-card" onclick="filterBySpecialStatus('next_sd')" style="cursor:pointer; border-color:rgba(255,171,0,0.4);" title="Click to view assets scheduled for Next Shutdown">
          <div class="n" style="color:#ffab00;">${d.next_sd_count || 0}</div>
          <div class="label">Next Shutdown (SD)</div>
        </div>
        <div class="stat-card highrisk">
          <div class="n">${d.high_risk_count}</div>
          <div class="label">High Risk (API 580)</div>
        </div>
        <div class="stat-card cui">
          <div class="n">${d.cui_count}</div>
          <div class="label">CUI Flagged Assets</div>
        </div>
        <div class="stat-card due30">
          <div class="n">${d.due_30_count}</div>
          <div class="label">Due in Next 30 Days</div>
        </div>
      `;
    }

    // 3. Overdue Aging Triage Cards
    const triageGrid = document.getElementById("triageGrid");
    if (triageGrid && d.aging_counts) {
      const ac = d.aging_counts;
      triageGrid.innerHTML = `
        <div class="triage-card c-0-30" onclick="filterByAging('0_30')">
          <div class="n">${ac["0_30"]}</div>
          <div class="lbl">0 – 30 Days Overdue</div>
        </div>
        <div class="triage-card c-31-90" onclick="filterByAging('31_90')">
          <div class="n">${ac["31_90"]}</div>
          <div class="lbl">31 – 90 Days Overdue</div>
        </div>
        <div class="triage-card c-91-180" onclick="filterByAging('91_180')">
          <div class="n">${ac["91_180"]}</div>
          <div class="lbl">91 – 180 Days Overdue</div>
        </div>
        <div class="triage-card c-180-plus" onclick="filterByAging('180_plus')">
          <div class="n">${ac["180_plus"]}</div>
          <div class="lbl">>180 Days (Critical)</div>
        </div>
        <div class="triage-card c-deferred" onclick="filterByDeferral('deferred')">
          <div class="n">${ac["approved_deferrals"]}</div>
          <div class="lbl">Approved MOC Deferrals</div>
        </div>
        <div class="triage-card c-180-plus" onclick="filterByAging('')" style="opacity:0.85;" title="Overdue assets whose due date is an import placeholder — excluded from aging buckets">
          <div class="n">${ac["estimated_date"] || 0}</div>
          <div class="lbl">Estimated-Date Items ⚠</div>
        </div>
      `;
    }

    // 4. Render 5x5 RBI Matrix
    if (d.rbi_matrix) {
      render5x5Matrix(d.rbi_matrix);
    }

    // 4b. Missing-due-date worklist
    const noDueBadge = document.getElementById("noDueDateBadge");
    if (noDueBadge) noDueBadge.textContent = d.no_due_date_count || 0;
    const noDueEl = document.getElementById("noDueDateList");
    if (noDueEl) {
      const items = d.no_due_date || [];
      noDueEl.innerHTML = items.length ? `<table class="assets"><thead><tr><th>Asset</th><th>Sheet</th><th>Risk</th><th>Last OSI</th></tr></thead><tbody>
        ${items.slice(0, 25).map(a => `<tr data-id="${a.id}" style="cursor:pointer;" onclick="openAssetDrawer(${a.id})">
          <td>${esc(a.name || a.tag || "—")}</td>
          <td>${esc(a.source_sheet || "")}</td>
          <td>${riskBadge(a)}</td>
          <td>${esc(a.date_osi_last || "Never inspected")}</td>
        </tr>`).join("")}</tbody></table>
        ${items.length > 25 ? `<div style="color:var(--text-faint); font-size:12px; padding:8px;">+ ${items.length - 25} more — filter Assets view to clear them</div>` : ""}`
        : `<div style="color:var(--text-faint); font-size:12.5px; padding:10px;">Every asset has a due date. 🎉</div>`;
    }

    // 4c. Audit trail + backups
    loadAuditTrail();
    loadBackups();

    // 5. High Risk List
    const highRiskEl = document.getElementById("highRiskList");
    if (highRiskEl) {
      highRiskEl.innerHTML = tableHtml(d.high_risk || [], true);
      bindRowClicks(highRiskEl);
      const countEl = document.getElementById("highRiskCountBadge");
      if (countEl) countEl.textContent = (d.high_risk || []).length;
    }

    // 6. Overdue List
    const overdueEl = document.getElementById("overdueList");
    if (overdueEl) {
      overdueEl.innerHTML = tableHtml(d.overdue || [], true);
      bindRowClicks(overdueEl);
      const countEl = document.getElementById("overdueCountBadge");
      if (countEl) countEl.textContent = (d.overdue || []).length;
    }

    // 7. Temporary Repairs Widget
    const tempRepairsWidgetEl = document.getElementById("tempRepairsWidgetList");
    if (tempRepairsWidgetEl) {
      const repairs = d.expired_repairs || [];
      const countEl = document.getElementById("tempRepairsCountBadge");
      if (countEl) countEl.textContent = repairs.length;
      if (!repairs.length) {
        tempRepairsWidgetEl.innerHTML = `<table class="assets"><tbody><tr class="empty-row"><td>No expired temporary repairs.</td></tr></tbody></table>`;
      } else {
        tempRepairsWidgetEl.innerHTML = `<table class="assets">
          <thead><tr><th>Area / Asset</th><th>Repaired Section</th><th>Contractor</th><th>Expiry Date</th></tr></thead>
          <tbody>${repairs.map(r => `<tr>
            <td><div style="font-weight:600; color:#fff;">${esc(r.asset_name || r.area || "")}</div><div class="tag-mono">${esc(r.facility_type || "")}</div></td>
            <td><div style="color:var(--accent); font-size:12px;">${esc(r.repaired_section || "")}</div></td>
            <td>${esc(r.repaired_by || "—")}</td>
            <td><span class="status-pill overdue">${esc(r.expiration_date || "Expired")}</span></td>
          </tr>`).join("")}</tbody>
        </table>`;
      }
    }

    // 8. Source Sheet Chips
    const bySheetEl = document.getElementById("bySheet");
    if (bySheetEl && d.sheets && d.by_sheet) {
      state.sheets = d.sheets;
      state.by_sheet = d.by_sheet;
      bySheetEl.innerHTML = d.sheets.map(s => {
        const info = d.by_sheet[s] || {total:0, overdue:0};
        return `<span class="chip" onclick="filterBySheet('${esc(s)}')">
          <b>${esc(s)}</b> — ${info.total} ${info.overdue ? `· <span class="od">${info.overdue} overdue</span>` : ""}
        </span>`;
      }).join("");

      const sel = document.getElementById("sheetFilter");
      if (sel) {
        const current = sel.value;
        sel.innerHTML = `<option value="">All Sheets (${d.sheets.length} Systems)</option>` + 
          d.sheets.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
        sel.value = current;
      }

      const ysel = document.getElementById("yearlySheetFilter");
      if (ysel) {
        const currentY = ysel.value;
        ysel.innerHTML = `<option value="">All Sheets (Global)</option>` + 
          d.sheets.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
        ysel.value = currentY;
      }
      
      renderSheetTabs(d.sheets, sel ? sel.value : "");
    }

  } catch (err) {
    showToast("Failed to load dashboard: " + err.message, true);
  }
}

function renderSheetTabs(sheets, activeSheet) {
  const bar = document.getElementById("sheetTabsBar");
  if (!bar) return;
  
  const allCount = state.assets?.length || 0;
  const isAllActive = !activeSheet;
  
  let html = `<button class="sheet-tab-btn ${isAllActive ? 'active' : ''}" onclick="selectSheetTab('')">
    📑 All Sheets <span class="badge-num">${allCount ? allCount.toLocaleString() : '1,288'}</span>
  </button>`;
  
  (sheets || state.sheets || []).forEach(s => {
    const isActive = activeSheet === s;
    const count = (state.by_sheet && state.by_sheet[s]) ? state.by_sheet[s].total : '';
    html += `<button class="sheet-tab-btn ${isActive ? 'active' : ''}" onclick="selectSheetTab('${esc(s)}')">
      📂 ${esc(s)} ${count ? `<span class="badge-num">${count}</span>` : ''}
    </button>`;
  });
  
  bar.innerHTML = html;
}

window.selectSheetTab = function(sheetName) {
  const sel = document.getElementById("sheetFilter");
  if (sel) sel.value = sheetName;
  state.page = 1;
  loadAssets();
  renderSheetTabs(state.sheets, sheetName);
};

window.filterBySheet = function(sheetName) {
  const sel = document.getElementById("sheetFilter");
  if (sel) sel.value = sheetName;
  switchView("assets");
  renderSheetTabs(state.sheets, sheetName);
};

window.filterByAging = function(bucket) {
  state.activeAgingBucket = bucket;
  const af = document.getElementById("agingFilter");
  if (af) af.value = bucket;
  switchView("assets");
};

window.filterByDeferral = function(status) {
  const df = document.getElementById("deferralFilter");
  if (df) df.value = status;
  switchView("assets");
};

window.filterByOverdueType = function(type) {
  const oo = document.getElementById("overdueOnly");
  if (oo) oo.checked = true;
  const otf = document.getElementById("overdueTypeFilter");
  if (otf) otf.value = type;
  const ssf = document.getElementById("specialStatusFilter");
  if (ssf) ssf.value = "";
  switchView("assets");
};

window.filterBySpecialStatus = function(st) {
  const oo = document.getElementById("overdueOnly");
  if (oo) oo.checked = false;
  const otf = document.getElementById("overdueTypeFilter");
  if (otf) otf.value = "";
  const ssf = document.getElementById("specialStatusFilter");
  if (ssf) ssf.value = st;
  switchView("assets");
};

// ---------------------------------------------------------------- Assets Inventory
let searchDebounce = null;
const searchBox = document.getElementById("searchBox");
const btnClearSearch = document.getElementById("btnClearSearch");

if (searchBox) {
  searchBox.addEventListener("input", () => {
    if (btnClearSearch) btnClearSearch.style.display = searchBox.value ? "block" : "none";
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => {
      state.page = 1;
      loadAssets();
    }, 180);
  });
}

if (btnClearSearch) {
  btnClearSearch.addEventListener("click", () => {
    searchBox.value = "";
    btnClearSearch.style.display = "none";
    state.page = 1;
    loadAssets();
  });
}

document.getElementById("sheetFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("riskFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("overdueTypeFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("specialStatusFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("agingFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("deferralFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("groupFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("overdueOnly")?.addEventListener("change", () => { state.page = 1; loadAssets(); });

document.getElementById("btnResetFilters")?.addEventListener("click", () => {
  if (searchBox) searchBox.value = "";
  if (btnClearSearch) btnClearSearch.style.display = "none";
  const sf = document.getElementById("sheetFilter"); if (sf) sf.value = "";
  const rf = document.getElementById("riskFilter"); if (rf) rf.value = "";
  const otf = document.getElementById("overdueTypeFilter"); if (otf) otf.value = "";
  const ssf = document.getElementById("specialStatusFilter"); if (ssf) ssf.value = "";
  const af = document.getElementById("agingFilter"); if (af) af.value = "";
  const df = document.getElementById("deferralFilter"); if (df) df.value = "";
  const gf = document.getElementById("groupFilter"); if (gf) gf.value = "";
  const oo = document.getElementById("overdueOnly"); if (oo) oo.checked = false;
  state.activeRbiCell = "";
  state.page = 1;
  loadAssets();
});

// Pagination controls
document.getElementById("pageSizeSelect")?.addEventListener("change", (e) => {
  state.pageSize = e.target.value === "all" ? "all" : parseInt(e.target.value);
  state.page = 1;
  renderAssetsTable();
});

document.getElementById("btnPrevPage")?.addEventListener("click", () => {
  if (state.page > 1) {
    state.page--;
    renderAssetsTable();
  }
});

document.getElementById("btnNextPage")?.addEventListener("click", () => {
  const totalPages = state.pageSize === "all" ? 1 : Math.ceil(state.filteredAssets.length / state.pageSize);
  if (state.page < totalPages) {
    state.page++;
    renderAssetsTable();
  }
});

async function loadAssets() {
  try {
    const q = searchBox ? searchBox.value.trim() : "";
    const sheet = document.getElementById("sheetFilter")?.value || "";
    const risk = document.getElementById("riskFilter")?.value || "";
    const overdueType = document.getElementById("overdueTypeFilter")?.value || "";
    const specialStatus = document.getElementById("specialStatusFilter")?.value || "";
    const aging = document.getElementById("agingFilter")?.value || "";
    const deferral = document.getElementById("deferralFilter")?.value || "";
    const overdue = document.getElementById("overdueOnly")?.checked ? "1" : "0";
    state.groupBy = document.getElementById("groupFilter")?.value || "";

    const params = new URLSearchParams({
      q, sheet, risk, overdue,
      overdue_type: overdueType,
      status: specialStatus,
      aging, deferral,
      rbi_cell: state.activeRbiCell
    });

    const rows = await api("/api/assets?" + params.toString());
    state.assets = rows;
    state.filteredAssets = rows;

    if (!state.sheets || !state.sheets.length) {
      api("/api/dashboard").then(d => {
        if (d.sheets) {
          state.sheets = d.sheets;
          state.by_sheet = d.by_sheet;
          renderSheetTabs(d.sheets, sheet);
        }
      }).catch(() => {});
    } else {
      renderSheetTabs(state.sheets, sheet);
    }

    renderAssetsTable();
  } catch (err) {
    showToast("Failed to load assets: " + err.message, true);
  }
}

function renderAssetsTable() {
  const rows = state.filteredAssets;
  const totalCount = rows.length;
  const countTextEl = document.getElementById("assetCountText");
  const pageIndicatorEl = document.getElementById("pageIndicator");
  const btnPrev = document.getElementById("btnPrevPage");
  const btnNext = document.getElementById("btnNextPage");

  let pageRows = rows;
  let totalPages = 1;

  if (state.pageSize !== "all") {
    totalPages = Math.max(1, Math.ceil(totalCount / state.pageSize));
    if (state.page > totalPages) state.page = totalPages;
    const startIdx = (state.page - 1) * state.pageSize;
    const endIdx = startIdx + state.pageSize;
    pageRows = rows.slice(startIdx, endIdx);

    if (countTextEl) {
      const shownStart = totalCount === 0 ? 0 : startIdx + 1;
      const shownEnd = Math.min(endIdx, totalCount);
      countTextEl.textContent = `Showing ${shownStart}–${shownEnd} of ${totalCount.toLocaleString()} assets`;
    }
  } else {
    if (countTextEl) {
      countTextEl.textContent = `Showing all ${totalCount.toLocaleString()} assets`;
    }
  }

  if (pageIndicatorEl) {
    pageIndicatorEl.textContent = `Page ${state.page} of ${totalPages}`;
  }
  if (btnPrev) btnPrev.disabled = state.page <= 1;
  if (btnNext) btnNext.disabled = state.page >= totalPages;

  const wrapEl = document.getElementById("assetsTableWrap");
  if (!wrapEl) return;

  if (!pageRows.length) {
    wrapEl.innerHTML = `<table class="assets"><tbody><tr class="empty-row"><td>No assets matched your search filters.</td></tr></tbody></table>`;
    return;
  }

  if (state.groupBy) {
    const groups = {};
    pageRows.forEach(a => {
      const key = a[state.groupBy] || "Unassigned";
      if (!groups[key]) groups[key] = [];
      groups[key].push(a);
    });

    let html = `<table class="assets">
      <thead><tr><th style="width:36px; text-align:center;"><input type="checkbox" id="selectAllCheckbox"></th><th>Asset & Risk</th><th>Source Sheet</th><th>Status</th><th>Pack # / Header</th><th>Plant & Location</th></tr></thead>
      <tbody>`;

    for (const [groupName, groupAssets] of Object.entries(groups)) {
      html += `<tr class="group-header-row"><td colspan="6">📂 ${esc(groupName)} (${groupAssets.length} assets)</td></tr>`;
      html += groupAssets.map(a => assetRowHtml(a, false)).join("");
    }
    html += `</tbody></table>`;
    wrapEl.innerHTML = html;
  } else {
    wrapEl.innerHTML = tableHtml(pageRows, false);
  }

  bindRowClicks(wrapEl);
}

// ---------------------------------------------------------------- Batch Action Bar & Modals
function updateBatchActionBar() {
  const bar = document.getElementById("batchActionBar");
  const badge = document.getElementById("selectedCountBadge");
  const count = state.selectedAssetIds.size;
  if (badge) badge.textContent = count;
  if (bar) bar.classList.toggle("open", count > 0);
}

document.getElementById("btnClearSelection")?.addEventListener("click", () => {
  state.selectedAssetIds.clear();
  document.querySelectorAll(".row-checkbox, #selectAllCheckbox").forEach(cb => cb.checked = false);
  updateBatchActionBar();
});

document.getElementById("btnExportSelected")?.addEventListener("click", () => {
  if (!state.selectedAssetIds.size) return;
  const selectedRows = state.assets.filter(a => state.selectedAssetIds.has(a.id));
  const csv = convertToCSV(selectedRows);
  downloadCSV(csv, `selected_assets_${selectedRows.length}.csv`);
});

function convertToCSV(arr) {
  if (!arr.length) return "";
  const keys = Object.keys(arr[0]).filter(k => k !== "extra");
  let out = keys.join(",") + "\n";
  arr.forEach(row => {
    out += keys.map(k => `"${String(row[k] ?? "").replace(/"/g, '""')}"`).join(",") + "\n";
  });
  return out;
}

function downloadCSV(csvText, filename) {
  const blob = new Blob([csvText], {type: "text/csv;charset=utf-8;"});
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// Bulk Modals
document.getElementById("btnOpenBulkLog")?.addEventListener("click", () => {
  const count = state.selectedAssetIds.size;
  if (!count) return;
  document.getElementById("bulkLogCountText").textContent = count;
  document.getElementById("bulkLogDate").value = new Date().toISOString().split("T")[0];
  document.getElementById("bulkLogModal").classList.add("open");
});

document.getElementById("btnOpenBulkDeferral")?.addEventListener("click", () => {
  const count = state.selectedAssetIds.size;
  if (!count) return;
  document.getElementById("bulkDeferralCountText").textContent = count;
  document.getElementById("bulkDeferralModal").classList.add("open");
});

window.closeBulkModals = function() {
  document.querySelectorAll(".modal-overlay").forEach(m => m.classList.remove("open"));
};

document.getElementById("btnSubmitBulkLog")?.addEventListener("click", async () => {
  const assetIds = Array.from(state.selectedAssetIds);
  const payload = {
    asset_ids: assetIds,
    insp_date: document.getElementById("bulkLogDate")?.value || null,
    insp_type: document.getElementById("bulkLogType")?.value || "OSI",
    inspector_name: document.getElementById("bulkLogInspector")?.value?.trim() || null,
    insp_method: document.getElementById("bulkLogMethod")?.value?.trim() || null,
    findings: document.getElementById("bulkLogFindings")?.value?.trim() || null,
    next_due_date: document.getElementById("bulkLogNextDue")?.value || null,
  };

  try {
    const res = await api("/api/assets/bulk-log", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload)
    });
    showToast(res.message || "Bulk inspection applied successfully!", false);
    closeBulkModals();
    state.selectedAssetIds.clear();
    updateBatchActionBar();
    refreshCurrentView();
  } catch (err) {
    showToast("Bulk log failed: " + err.message, true);
  }
});

document.getElementById("btnSubmitBulkDeferral")?.addEventListener("click", async () => {
  const assetIds = Array.from(state.selectedAssetIds);
  const payload = {
    asset_ids: assetIds,
    deferral_status: document.getElementById("bulkDefStatus")?.value || "Approved",
    deferral_moc_no: document.getElementById("bulkDefMoc")?.value?.trim() || null,
    deferral_approver: document.getElementById("bulkDefApprover")?.value?.trim() || null,
    deferral_expiry: document.getElementById("bulkDefExpiry")?.value || null,
    deferral_reason: document.getElementById("bulkDefReason")?.value?.trim() || null,
    deferral_mitigation: document.getElementById("bulkDefMitigation")?.value?.trim() || null,
  };

  try {
    const res = await api("/api/assets/bulk-deferral", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload)
    });
    showToast(res.message || "Bulk deferral applied successfully!", false);
    closeBulkModals();
    state.selectedAssetIds.clear();
    updateBatchActionBar();
    refreshCurrentView();
  } catch (err) {
    showToast("Bulk deferral failed: " + err.message, true);
  }
});

// ---------------------------------------------------------------- Slide-Over Asset Drawer & Calculator
const overlay = document.getElementById("drawerOverlay");
document.getElementById("drawerClose")?.addEventListener("click", closeDrawer);
overlay?.addEventListener("click", e => { if (e.target === overlay) closeDrawer(); });

function closeDrawer() {
  if (overlay) overlay.classList.remove("open");
  state.editingAsset = null;
}

document.addEventListener("keydown", e => {
  if (e.key === "Escape") {
    closeDrawer();
    closeBulkModals();
  }
});

const CORE_FIELDS = [
  ["name", "Asset Name"], ["field", "Pack # / Header Section"], ["location", "Location / Area"],
  ["tag", "Equipment Tag #"], ["asset_number", "Asset ID Number"], ["plant", "Plant / Facility"],
  ["unit_name", "Unit Name"], ["source_sheet", "Source Sheet"], ["in_service", "In Service Status"],
  ["insulation", "Insulation (CUI Factor)"],
];

const ENVELOPE_FIELDS = [
  ["fluid_service", "Fluid Service"], ["material_spec", "Material Specification"],
  ["design_pressure", "Design Pressure (psi/bar)"], ["operating_pressure", "Operating Pressure"],
  ["design_temp", "Design Temp (°C)"], ["operating_temp", "Operating Temp (°C)"],
  ["nominal_thickness", "Nominal Thickness (mm)"], ["t_min", "Min Required t_min (mm)"],
];

const DATE_FIELDS = [
  ["date_osi_last", "Last OSI Inspection Date"], ["date_osi_next", "Next OSI Due Date"],
  ["date_internal_last", "Last Internal Inspection Date"], ["date_internal_next", "Next Internal Due Date"],
  ["status_osi_next", "OSI Status / Condition"], ["status_internal_next", "Internal Status / Condition"],
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
  try {
    const a = await api("/api/assets/" + id);
    state.editingAsset = a;
    
    const titleEl = document.getElementById("drawerTitle");
    if (titleEl) titleEl.innerHTML = `${esc(a.name || "Asset")} ${riskBadge(a)}`;

    const subEl = document.getElementById("drawerSubtitle");
    if (subEl) subEl.textContent = `ID #${a.id} · Tag: ${a.tag || a.sn || "—"} · Sheet: ${a.source_sheet || "—"}`;

    const extraRows = Object.entries(a.extra || {}).filter(([k]) => !k.endsWith("(raw)"));

    const bodyEl = document.getElementById("drawerBody");
    if (!bodyEl) return;

    bodyEl.innerHTML = `
      <div class="drawer-tabs">
        <button class="drawer-tab active" id="tabbtn-specs" onclick="switchDrawerTab('specs')">📋 Specs & Envelope</button>
        <button class="drawer-tab" id="tabbtn-risk" onclick="switchDrawerTab('risk')">🛡️ Risk & Deferrals</button>
        <button class="drawer-tab" id="tabbtn-logs" onclick="switchDrawerTab('logs')">🔍 NDT & Calculator (${(a.log || []).length})</button>
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

      <!-- TAB 2: RISK & DEFERRALS -->
      <div id="dtab-risk" class="dtab-content" style="display:none;">
        <div class="section-title">API 580 RBI Risk & Damage Mechanisms</div>
        <div class="field-grid">
          <div class="field"><label>RBI Risk Level</label><select data-field="risk_category">
            <option value="HIGH" ${a.risk_category==="HIGH"?"selected":""}>HIGH (Priority Tier 1)</option>
            <option value="MEDIUM" ${a.risk_category==="MEDIUM"?"selected":""}>MEDIUM (Tier 2)</option>
            <option value="LOW" ${a.risk_category==="LOW"?"selected":""}>LOW (Tier 3)</option>
          </select></div>
          <div class="field"><label>RBI Matrix 5x5 Cell</label><input value="${esc(a.rbi_matrix_cell || '')}" readonly></div>
        </div>

        <div class="field-grid">
          <div class="field"><label>Corrosion Rate (mm/yr)</label><input data-field="corrosion_rate" value="${esc(a.corrosion_rate || '')}"></div>
          <div class="field"><label>Estimated Remaining Life (yrs)</label><input data-field="remaining_life" value="${esc(a.remaining_life || '')}"></div>
        </div>

        <div class="section-title">Scheduled Inspection Due Dates</div>
        <div class="field-grid">
          ${DATE_FIELDS.map(([k,l]) => fieldInput(k,l,a[k],"text")).join("")}
        </div>

        <!-- Formal Deferral Governance Form -->
        <div class="section-title">🛡️ Formal Inspection Deferral & MOC Concession</div>
        <div class="field-grid">
          <div class="field"><label>Deferral Status</label><select data-field="deferral_status">
            <option value="None" ${a.deferral_status==="None"||!a.deferral_status?"selected":""}>None (No Active Deferral)</option>
            <option value="Approved" ${a.deferral_status==="Approved"?"selected":""}>Approved (Valid MOC Concession)</option>
            <option value="Under Review" ${a.deferral_status==="Under Review"?"selected":""}>Under Review (MOC In-Progress)</option>
            <option value="Expired" ${a.deferral_status==="Expired"?"selected":""}>Expired Concession</option>
          </select></div>
          <div class="field"><label>MOC / Concession Reference #</label><input data-field="deferral_moc_no" value="${esc(a.deferral_moc_no || '')}"></div>
        </div>

        <div class="field-grid">
          <div class="field"><label>Approver Name & Authority</label><input data-field="deferral_approver" value="${esc(a.deferral_approver || '')}"></div>
          <div class="field"><label>Deferral Expiration Date</label><input data-field="deferral_expiry" type="date" value="${esc(a.deferral_expiry || '')}"></div>
        </div>

        <div class="field-grid wide">
          <div class="field"><label>Technical Justification</label><textarea data-field="deferral_reason">${esc(a.deferral_reason || '')}</textarea></div>
          <div class="field"><label>Compensating Mitigating Measures</label><textarea data-field="deferral_mitigation">${esc(a.deferral_mitigation || '')}</textarea></div>
          <div class="field"><label>General Integrity Remarks</label><textarea data-field="remarks">${esc(a.remarks || '')}</textarea></div>
        </div>

        ${extraRows.length ? `
          <div class="section-title">Original Sheet Column Metadata</div>
          <div class="field-grid wide">
            ${extraRows.map(([k,v]) => `<div class="field"><label>${esc(k)}</label>
              <input value="${esc(v)}" readonly></div>`).join("")}
          </div>` : ""}
      </div>

      <!-- TAB 3: LOGS & CALCULATOR -->
      <div id="dtab-logs" class="dtab-content" style="display:none;">
        
        <!-- API 510/570 Auto-Calculator Widget -->
        <div class="calc-card">
          <div style="font-weight:700; color:#fff; font-size:13px; display:flex; justify-content:space-between;">
            <span>⚡ API 510/570 Remaining Life & Next Due Calculator</span>
          </div>
          <div class="field-grid" style="margin-top:10px;">
            <div class="field"><label>Nominal/Prev Thickness (mm)</label><input id="calcTPrev" value="${esc(a.nominal_thickness || '8.0')}"></div>
            <div class="field"><label>Actual Measured t_act (mm)</label><input id="calcTAct" placeholder="e.g. 6.8"></div>
          </div>
          <div class="field-grid">
            <div class="field"><label>Min Required t_min (mm)</label><input id="calcTMin" value="${esc(a.t_min || '3.5')}"></div>
            <div class="field"><label>Years Between Inspections</label><input id="calcYears" value="5.0"></div>
          </div>
          <div style="margin-top:8px;">
            <button class="btn ghost small" onclick="computeRemainingLife()">Calculate RL & Next Due</button>
            <button class="btn primary small" id="btnApplyCalc" style="display:none;" onclick="applyCalcToForm()">Apply to Form</button>
          </div>
          <div class="calc-results-grid" id="calcResultsGrid" style="display:none;">
            <div class="calc-res-item"><div class="calc-res-val" id="resCR">0.000</div><div class="calc-res-lbl">CR (mm/yr)</div></div>
            <div class="calc-res-item"><div class="calc-res-val" id="resRL">0.0</div><div class="calc-res-lbl">Remaining Life (yrs)</div></div>
            <div class="calc-res-item"><div class="calc-res-val" id="resInterval">0.0</div><div class="calc-res-lbl">Half-Life Int (yrs)</div></div>
            <div class="calc-res-item"><div class="calc-res-val" id="resDue" style="font-size:13px;">—</div><div class="calc-res-lbl">Suggested Due Date</div></div>
          </div>
        </div>

        <div class="section-title">Record New Inspection Activity</div>
        <div class="field-grid">
          <div class="field"><label>Inspection Date</label><input id="logDate" type="date" value="${new Date().toISOString().split('T')[0]}"></div>
          <div class="field"><label>Inspection Type</label>
            <select id="logType">
              <option value="OSI">OSI (On-Stream Inspection)</option>
              <option value="Internal">Internal Inspection</option>
              <option value="OSI-ADV">OSI Advanced NDT (PAUT/TOFD)</option>
              <option value="NDT UT Thickness">NDT UT Wall Thickness</option>
              <option value="Visual / CUI">Visual / CUI Surveillance</option>
            </select>
          </div>
        </div>
        <div class="field-grid">
          <div class="field"><label>Inspector Name & Cert #</label><input id="logInspector" placeholder="e.g. M. Ahmed (API 570 #4421)"></div>
          <div class="field"><label>NDT Technique</label><input id="logMethod" placeholder="e.g. Digital UT, MPT"></div>
        </div>
        <div class="field-grid">
          <div class="field"><label>Actual Measured Min Thickness t_act (mm)</label><input id="logTActual" placeholder="e.g. 6.8 mm"></div>
          <div class="field"><label>Corrective Action Required?</label><input id="logAction" placeholder="e.g. None / Turnaround Scope"></div>
        </div>
        <div class="field-grid wide">
          <div class="field"><label>Findings & Observations</label><textarea id="logFindings" placeholder="Document wall thinning, corrosion patterns, coating condition, recommendations…"></textarea></div>
        </div>
        <div class="field-grid">
          <div class="field"><label>Next Scheduled Due Date (YYYY-MM-DD)</label><input id="logNextDue" type="date"></div>
        </div>
        <div class="drawer-actions" style="margin-top:12px; margin-bottom:24px;">
          <button id="btnAddLog" class="btn primary">💾 Save Inspection Record</button>
        </div>

        ${corrosionTrendHtml(a)}

        ${a.log && a.log.length ? `
          <div class="section-title">Historical Inspection & NDT Log</div>
          ${a.log.map(l => `<div class="log-entry">
            <div class="meta">${esc(l.insp_date || "Date N/A")} · ${esc(l.insp_type || "Inspection")}${l.insp_method ? " · " + esc(l.insp_method) : ""}${l.next_due_date ? " · Next Due: " + esc(l.next_due_date) : ""}</div>
            ${l.inspector_name ? `<div style="font-weight:600; font-size:12px; color:var(--accent);">Inspector: ${esc(l.inspector_name)}</div>` : ""}
            ${l.t_actual ? `<div style="font-size:12px; color:var(--text-muted);">Measured Wall Thickness: <b>${esc(l.t_actual)}</b></div>` : ""}
            ${l.action_required ? `<div style="font-size:12px; color:var(--amber);">Action: ${esc(l.action_required)}</div>` : ""}
            <div style="margin-top:6px; color:#fff; font-size:12.5px;">${esc(l.findings || "")}</div>
            <button class="btn danger small" style="position:absolute; top:10px; right:10px;" onclick="deleteLog(${a.id}, ${l.id})">Delete</button>
          </div>`).join("")}
        ` : `<div style="color:var(--text-faint); font-size:12.5px; margin-top:8px;">No historical inspection logs recorded yet.</div>`}
      </div>

      <div class="drawer-actions" style="margin-top:24px; border-top:1px solid var(--border); padding-top:16px;">
        <button id="btnSaveAsset" class="btn primary">💾 Save Changes</button>
        <button id="btnArchiveAsset" class="btn danger">${a.archived ? "♻️ Restore Asset" : "📦 Archive Asset"}</button>
      </div>
    `;

    document.getElementById("btnSaveAsset")?.addEventListener("click", () => saveAsset(id));
    document.getElementById("btnArchiveAsset")?.addEventListener("click", () => toggleArchive(id, !a.archived));
    document.getElementById("btnAddLog")?.addEventListener("click", () => addLog(id));

    if (overlay) overlay.classList.add("open");
  } catch (err) {
    showToast("Failed to open asset details: " + err.message, true);
  }
}

// ---------------------------------------------------------------- Corrosion Trend
function corrosionTrendHtml(asset) {
  const readings = (asset.log || [])
    .filter(l => l.t_actual && parseFloat(l.t_actual) > 0)
    .map(l => ({date: l.insp_date, t: parseFloat(l.t_actual)}))
    .filter(r => r.date)
    .sort((a, b) => a.date.localeCompare(b.date));

  if (readings.length < 2) return "";

  const first = readings[0], last = readings[readings.length - 1];
  const days = (new Date(last.date) - new Date(first.date)) / 86400000;
  const loss = first.t - last.t;
  const years = days / 365.25;
  const cr = years > 0 ? loss / years : 0;
  const tMin = parseFloat(asset.t_min) || 0;
  const rl = cr > 0 && tMin > 0 ? (last.t - tMin) / cr : (cr <= 0 ? 99 : 0);

  let verdict, cls;
  if (cr <= 0)      { verdict = "No measurable loss — establish a longer baseline before setting intervals"; cls = "var(--text-muted);"; }
  else if (cr < 0.1)  { verdict = "Low corrosion rate"; cls = "var(--ok, #4ade80);"; }
  else if (cr < 0.25) { verdict = "Moderate corrosion rate — monitor"; cls = "var(--amber);"; }
  else              { verdict = "HIGH corrosion rate — shorten intervals"; cls = "var(--danger);"; }

  return `
    <div class="section-title">📉 Corrosion Trend (from logged thickness readings)</div>
    <div class="calc-card" style="padding:12px 14px;">
      <div class="field-grid">
        <div class="calc-res-item"><div class="calc-res-val" style="font-size:16px;">${cr > 0 ? cr.toFixed(3) : "0"}</div><div class="calc-res-lbl">Avg CR (mm/yr)</div></div>
        <div class="calc-res-item"><div class="calc-res-val" style="font-size:16px;">${loss > 0 ? loss.toFixed(2) : "0"}</div><div class="calc-res-lbl">Loss (mm)</div></div>
        <div class="calc-res-item"><div class="calc-res-val" style="font-size:16px;">${years > 0 ? years.toFixed(1) : "—"}</div><div class="calc-res-lbl">Baseline (yrs)</div></div>
        <div class="calc-res-item"><div class="calc-res-val" style="font-size:16px;">${rl >= 99 ? "99+" : rl.toFixed(1)}</div><div class="calc-res-lbl">Est. Remaining Life (yrs)</div></div>
      </div>
      <div style="margin-top:8px; font-size:12px; color:${cls};">${verdict} · ${readings.length} readings, ${esc(first.date)} → ${esc(last.date)}</div>
      ${readings.map(r => `<div style="font-size:11.5px; color:var(--text-muted); margin-top:2px;">• ${esc(r.date)}: t = ${r.t} mm</div>`).join("")}
    </div>`;
}

// ---------------------------------------------------------------- Calculator Functions
let lastCalcResult = null;

window.computeRemainingLife = async function() {
  const t_prev = parseFloat(document.getElementById("calcTPrev")?.value || 0);
  const t_act = parseFloat(document.getElementById("calcTAct")?.value || 0);
  const t_min = parseFloat(document.getElementById("calcTMin")?.value || 0);
  const years = parseFloat(document.getElementById("calcYears")?.value || 1);

  if (!t_act || t_act <= 0) {
    showToast("Please enter a valid actual measured thickness (t_act).", true);
    return;
  }

  try {
    const res = await api("/api/calc/remaining-life", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({t_prev, t_act, t_min, years_between: years})
    });

    if (!res.success) throw new Error(res.error || "Calculation failed");
    lastCalcResult = res;

    document.getElementById("resCR").textContent = `${res.corrosion_rate_mm_yr}`;
    document.getElementById("resRL").textContent = `${res.remaining_life_years}`;
    document.getElementById("resInterval").textContent = `${res.half_life_interval_years}`;
    document.getElementById("resDue").textContent = res.suggested_next_due_date || "—";

    document.getElementById("calcResultsGrid").style.display = "grid";
    document.getElementById("btnApplyCalc").style.display = "inline-flex";
  } catch (err) {
    showToast("Calculation error: " + err.message, true);
  }
};

window.applyCalcToForm = function() {
  if (!lastCalcResult) return;
  const logT = document.getElementById("logTActual");
  const logDue = document.getElementById("logNextDue");
  if (logT) logT.value = `${lastCalcResult.t_act} mm`;
  if (logDue && lastCalcResult.suggested_next_due_date) logDue.value = lastCalcResult.suggested_next_due_date;
  showToast("Calculated thickness and next due date applied to form!", false);
};

window.switchDrawerTab = function(tabName) {
  document.querySelectorAll(".drawer-tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".dtab-content").forEach(c => c.style.display = "none");
  
  const btn = document.getElementById("tabbtn-" + tabName);
  if (btn) btn.classList.add("active");
  const tabContent = document.getElementById("dtab-" + tabName);
  if (tabContent) tabContent.style.display = "block";
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
    showToast("Asset changes saved successfully!", false);
  } catch (err) {
    showToast("Save failed: " + err.message, true);
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
    showToast(`Asset ${archived ? "archived" : "restored"}`, false);
  } catch (err) {
    showToast("Archive toggle failed: " + err.message, true);
  }
}

async function addLog(id) {
  const dateVal = document.getElementById("logDate")?.value;
  const nextDueVal = document.getElementById("logNextDue")?.value;

  const payload = {
    insp_date: dateVal ? dateVal.trim() : null,
    insp_type: document.getElementById("logType")?.value || "OSI",
    inspector_name: document.getElementById("logInspector")?.value?.trim() || null,
    insp_method: document.getElementById("logMethod")?.value?.trim() || null,
    t_actual: document.getElementById("logTActual")?.value?.trim() || null,
    action_required: document.getElementById("logAction")?.value?.trim() || null,
    findings: document.getElementById("logFindings")?.value?.trim() || null,
    next_due_date: nextDueVal ? nextDueVal.trim() : null,
  };

  try {
    await api("/api/assets/" + id + "/log", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload),
    });
    showToast("Inspection record added successfully!", false);
    openAssetDrawer(id);
    refreshCurrentView();
  } catch (err) {
    showToast("Failed to add inspection log: " + err.message, true);
  }
}

window.deleteLog = async function(assetId, logId) {
  if (!confirm("Are you sure you want to delete this inspection record?")) return;
  try {
    await api("/api/logs/" + logId, { method: "DELETE" });
    showToast("Inspection record removed", false);
    openAssetDrawer(assetId);
    refreshCurrentView();
  } catch (err) {
    showToast("Failed to delete log: " + err.message, true);
  }
};

function refreshCurrentView() {
  if (state.view === "dashboard") loadDashboard();
  if (state.view === "assets") loadAssets();
  if (state.view === "critical") loadCriticalAssets();
  if (state.view === "temprepairs") loadTempRepairs();
  if (state.view === "yearly") loadYearlyPlan();
}

// ---------------------------------------------------------------- Critical Turnaround Replacements
document.getElementById("criticalSearch")?.addEventListener("input", loadCriticalAssets);
document.getElementById("criticalDoneFilter")?.addEventListener("change", loadCriticalAssets);

async function loadCriticalAssets() {
  try {
    const q = (document.getElementById("criticalSearch")?.value || "").toLowerCase().trim();
    const doneFilter = document.getElementById("criticalDoneFilter")?.value || "";
    let rows = await api("/api/critical_assets");
    
    if (q) {
      rows = rows.filter(r => (r.item_description || "").toLowerCase().includes(q) ||
                              (r.replacement_scope || "").toLowerCase().includes(q) ||
                              (r.pack_no || "").toLowerCase().includes(q) ||
                              (r.category_section || "").toLowerCase().includes(q) ||
                              (r.remarks || "").toLowerCase().includes(q));
    }
    if (doneFilter) {
      rows = rows.filter(r => (r.replacement_done || "").toLowerCase() === doneFilter.toLowerCase());
    }

    const countText = document.getElementById("criticalCountText");
    if (countText) countText.textContent = `${rows.length} critical scope items`;

    const el = document.getElementById("criticalTableWrap");
    if (!el) return;

    if (!rows.length) {
      el.innerHTML = `<table class="assets"><tbody><tr class="empty-row"><td>No matching critical replacement records.</td></tr></tbody></table>`;
      return;
    }

    el.innerHTML = `<table class="assets">
      <thead>
        <tr><th>SN</th><th>Pack #</th><th>Item Description</th><th>Category Section</th><th>Replacement Scope</th><th>Turnaround Status</th><th>Remarks</th></tr>
      </thead>
      <tbody>${rows.map(c => `<tr>
        <td><span class="tag-mono">${esc(c.sn || "—")}</span></td>
        <td><b style="color:var(--accent); font-family:var(--font-mono);">${esc(c.pack_no || "—")}</b></td>
        <td><div style="font-weight:600; color:#fff;">${esc(c.item_description || "—")}</div></td>
        <td><span class="tag-mono">${esc(c.category_section || "")}</span></td>
        <td><div style="max-width:320px; font-size:12.5px; color:var(--text-muted);">${esc(c.replacement_scope || "")}</div></td>
        <td>
          <button class="btn ${c.replacement_done === "Yes" ? "ghost" : "primary"} small" onclick="toggleCriticalDone(event, ${c.id}, '${c.replacement_done === "Yes" ? "No" : "Yes"}')">
            ${c.replacement_done === "Yes" ? "✅ Done (Yes)" : "⏳ Pending (No)"}
          </button>
        </td>
        <td><div style="font-size:12px; color:var(--text-faint);">${esc(c.remarks || "")} ${c.plant_remarks ? " | " + esc(c.plant_remarks) : ""}</div></td>
      </tr>`).join("")}</tbody>
    </table>`;
  } catch (err) {
    showToast("Failed to load critical turnaround assets: " + err.message, true);
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
    showToast(`Turnaround scope status updated to: ${doneStatus === "Yes" ? "Completed" : "Pending"}`, false);
  } catch (err) {
    showToast("Failed to update status: " + err.message, true);
  }
};

// ---------------------------------------------------------------- Temporary Repairs View
document.getElementById("tempRepairSearch")?.addEventListener("input", loadTempRepairs);
document.getElementById("tempRepairStatusFilter")?.addEventListener("change", loadTempRepairs);

async function loadTempRepairs() {
  try {
    const q = (document.getElementById("tempRepairSearch")?.value || "").toLowerCase().trim();
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
      if (statusFilter.toLowerCase() === "expired") {
        rows = rows.filter(r => (r.expiration_status || "").toLowerCase() === "expired" || (r.expiration_date && r.expiration_date < new Date().toISOString().split("T")[0]));
      } else {
        rows = rows.filter(r => (r.expiration_status || "").toLowerCase() === statusFilter.toLowerCase());
      }
    }

    const countText = document.getElementById("tempRepairCountText");
    if (countText) countText.textContent = `${rows.length} temporary repairs`;

    const el = document.getElementById("tempRepairsTableWrap");
    if (!el) return;

    if (!rows.length) {
      el.innerHTML = `<table class="assets"><tbody><tr class="empty-row"><td>No matching temporary repairs.</td></tr></tbody></table>`;
      return;
    }

    el.innerHTML = `<table class="assets">
      <thead>
        <tr><th>Facility / Area</th><th>Asset & Section</th><th>Contractor (Type)</th><th>Orig Date</th><th>Life (Yrs)</th><th>Expiry Date</th><th>Hardness (HB)</th><th>Status</th></tr>
      </thead>
      <tbody>${rows.map(r => {
        const isExp = (r.expiration_status || "").toLowerCase() === "expired" || (r.expiration_date && r.expiration_date < new Date().toISOString().split("T")[0]);
        const pillCls = (r.expiration_status || "").toLowerCase() === "replaced" ? "ok" : (isExp ? "overdue" : "soon");
        return `<tr>
          <td><div style="font-weight:700; color:#fff;">${esc(r.area || "")}</div><div class="tag-mono">${esc(r.facility_type || "")}</div></td>
          <td><div style="font-weight:600; color:#fff;">${esc(r.asset_name || "")}</div><div style="font-size:12px; color:var(--accent);">${esc(r.repaired_section || "")}</div></td>
          <td>${esc(r.repaired_by || "—")}</td>
          <td><span class="tag-mono">${esc(r.original_repair_date || "—")}</span></td>
          <td><span class="tag-mono">${esc(r.repair_life_years || "—")}</span></td>
          <td><b style="font-family:var(--font-mono); color:#fff;">${esc(r.expiration_date || "—")}</b></td>
          <td><span class="tag-mono">${esc(r.hardness_hb || "—")}</span></td>
          <td><span class="status-pill ${pillCls}">${esc(r.expiration_status || (isExp ? "Expired" : "Active"))}</span></td>
        </tr>`;
      }).join("")}</tbody>
    </table>`;
  } catch (err) {
    showToast("Failed to load temporary repairs: " + err.message, true);
  }
}

// ---------------------------------------------------------------- Yearly Plan
document.getElementById("btnLoadYear")?.addEventListener("click", loadYearlyPlan);
document.getElementById("yearlySheetFilter")?.addEventListener("change", loadYearlyPlan);
document.getElementById("yearlyTypeFilter")?.addEventListener("change", loadYearlyPlan);
document.getElementById("carryOverFromFilter")?.addEventListener("change", loadYearlyPlan);
document.getElementById("btnPrintYear")?.addEventListener("click", () => window.print());
document.getElementById("btnExportYearlyExcel")?.addEventListener("click", () => {
  const y = document.getElementById("yearInput")?.value || new Date().getFullYear();
  const sheet = document.getElementById("yearlySheetFilter")?.value || "";
  const type = document.getElementById("yearlyTypeFilter")?.value || "all";
  const carryFrom = document.getElementById("carryOverFromFilter")?.value || "";
  const params = new URLSearchParams({ year: y, sheet, type });
  if (carryFrom) params.append("carry_over_from", carryFrom);
  window.location.href = `/api/yearly_plan/export.xlsx?` + params.toString();
});

async function loadYearlyPlan() {
  try {
    const yearInput = document.getElementById("yearInput");
    const year = yearInput ? yearInput.value : new Date().getFullYear();
    const sheetFilter = document.getElementById("yearlySheetFilter");
    const selectedSheet = sheetFilter ? sheetFilter.value : "";
    const typeFilter = document.getElementById("yearlyTypeFilter");
    const scopeType = typeFilter ? typeFilter.value : "all";
    const carryFilter = document.getElementById("carryOverFromFilter");
    const carryFrom = carryFilter ? carryFilter.value : "";

    // Ensure sheet options are populated if empty
    if (sheetFilter && (!sheetFilter.options || sheetFilter.options.length <= 1) && state.sheets && state.sheets.length) {
      sheetFilter.innerHTML = `<option value="">All Sheets (Global)</option>` + 
        state.sheets.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
      sheetFilter.value = selectedSheet;
    }
    
    const params = new URLSearchParams({
      year,
      type: scopeType,
    });
    if (selectedSheet) {
      params.append("sheet", selectedSheet);
    }
    if (carryFrom) {
      params.append("carry_over_from", carryFrom);
    }
    
    const rows = await api(`/api/yearly_plan?` + params.toString());
    
    const metaEl = document.getElementById("yearlyMeta");
    if (metaEl) {
      const scopeLabel = scopeType === "all" ? "All Scopes" : (scopeType === "internal" ? "Internal (Shutdown) Only" : "OSI (On-Stream) Only");
      const sheetLabel = selectedSheet ? `Sheet: ${selectedSheet}` : "All Sheets (Global)";
      const carryLabel = carryFrom ? `Carry-over ≥ ${carryFrom}` : "All Overdue Carried";
      metaEl.textContent = `${rows.length} inspection event(s) scheduled for ${year} (${sheetLabel} · ${scopeLabel} · ${carryLabel})`;
    }
    
    const el = document.getElementById("yearlyTableWrap");
    if (el) {
      el.innerHTML = tableHtml(rows, false);
      bindRowClicks(el);
    }
  } catch (err) {
    showToast("Failed to load yearly plan: " + err.message, true);
  }
}

// ---------------------------------------------------------------- Export & Re-import
document.getElementById("btnExport")?.addEventListener("click", () => {
  const sheet = document.getElementById("sheetFilter")?.value || "";
  const params = new URLSearchParams(sheet ? {sheet} : {});
  window.location.href = "/api/export.csv?" + params.toString();
});

window.closeAllModals = function() {
  document.querySelectorAll(".modal-overlay").forEach(m => m.classList.remove("open"));
  const inp = document.getElementById("inputClearConfirmation");
  if (inp) inp.value = "";
  const btn = document.getElementById("btnExecuteClear");
  if (btn) btn.disabled = true;
};

// ---------------------------------------------------------------- Clear All Database Data
const btnClearDb = document.getElementById("btnClearDatabase");
const inputClearConfirm = document.getElementById("inputClearConfirmation");
const btnExecuteClear = document.getElementById("btnExecuteClear");

if (btnClearDb) {
  btnClearDb.addEventListener("click", () => {
    if (inputClearConfirm) inputClearConfirm.value = "";
    if (btnExecuteClear) btnExecuteClear.disabled = true;
    document.getElementById("clearDbModal")?.classList.add("open");
  });
}

if (inputClearConfirm && btnExecuteClear) {
  inputClearConfirm.addEventListener("input", () => {
    btnExecuteClear.disabled = inputClearConfirm.value.trim().toUpperCase() !== "CLEAR";
  });
}

if (btnExecuteClear) {
  btnExecuteClear.addEventListener("click", async () => {
    const origText = btnExecuteClear.innerHTML;
    btnExecuteClear.disabled = true;
    btnExecuteClear.innerHTML = `<span>⏳</span> Wiping Database…`;

    try {
      const res = await api("/api/database/clear", { method: "POST" });
      showToast(res.message || "Database cleared cleanly!", false);
      closeAllModals();
      refreshCurrentView();
    } catch (err) {
      showToast("Failed to clear database: " + err.message, true);
    } finally {
      btnExecuteClear.disabled = false;
      btnExecuteClear.innerHTML = origText;
    }
  });
}

// ---------------------------------------------------------------- Import Excel Modal & Workflow
let selectedImportFile = null;
const picker = document.getElementById("excelFilePicker");
const btnReimport = document.getElementById("btnReimport");
const btnConfirmImport = document.getElementById("btnConfirmImport");

if (btnReimport && picker) {
  btnReimport.addEventListener("click", () => {
    picker.value = "";
    picker.click();
  });

  picker.addEventListener("change", () => {
    const file = picker.files[0];
    if (!file) return;
    selectedImportFile = file;

    const fnEl = document.getElementById("importFileName");
    if (fnEl) fnEl.textContent = file.name;

    document.getElementById("reimportModal")?.classList.add("open");
  });
}

if (btnConfirmImport) {
  btnConfirmImport.addEventListener("click", async () => {
    if (!selectedImportFile) return;

    const mode = document.querySelector('input[name="importMode"]:checked')?.value || "clean";
    const cleanWipe = mode === "clean";

    const origText = btnConfirmImport.innerHTML;
    btnConfirmImport.disabled = true;
    btnConfirmImport.innerHTML = `<span>⏳</span> Ingesting Workbook…`;

    try {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const base64Data = e.target.result.split(",")[1];
          const res = await api("/api/reimport_file", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({
              filename: selectedImportFile.name,
              filedata: base64Data,
              clean_wipe: cleanWipe
            }),
          });
          showToast(res.message || "Workbook imported successfully!", false);
          closeAllModals();
          selectedImportFile = null;
          refreshCurrentView();
        } catch (err) {
          showToast("Import failed: " + err.message, true);
        } finally {
          btnConfirmImport.disabled = false;
          btnConfirmImport.innerHTML = origText;
        }
      };
      reader.onerror = () => {
        showToast("Failed to read selected file.", true);
        btnConfirmImport.disabled = false;
        btnConfirmImport.innerHTML = origText;
      };
      reader.readAsDataURL(selectedImportFile);
    } catch (err) {
      showToast("Error preparing file upload: " + err.message, true);
      btnConfirmImport.disabled = false;
      btnConfirmImport.innerHTML = origText;
    }
  });
}


// ---------------------------------------------------------------- Create New Asset
document.getElementById("btnNewAsset")?.addEventListener("click", () => {
  state.editingAsset = null;
  const titleEl = document.getElementById("drawerTitle");
  if (titleEl) titleEl.textContent = "Create New Asset";
  const subEl = document.getElementById("drawerSubtitle");
  if (subEl) subEl.textContent = "Manual equipment asset creation";

  const bodyEl = document.getElementById("drawerBody");
  if (!bodyEl) return;

  bodyEl.innerHTML = `
    <div class="section-title">Identification & Location</div>
    <div class="field-grid">
      ${CORE_FIELDS.map(([k,l]) => fieldInput(k,l,k==="source_sheet"?"Manual Entry":"")).join("")}
    </div>
    <div class="section-title">Design & Operating Envelope</div>
    <div class="field-grid">
      ${ENVELOPE_FIELDS.map(([k,l]) => fieldInput(k,l,"")).join("")}
    </div>
    <div class="section-title">Initial Scheduled Due Dates</div>
    <div class="field-grid">
      ${DATE_FIELDS.map(([k,l]) => fieldInput(k,l,"","date")).join("")}
    </div>
    <div class="field-grid wide">
      <div class="field">
        <label>Remarks & Integrity Recommendations</label>
        <textarea data-field="remarks"></textarea>
      </div>
    </div>
    <div class="drawer-actions">
      <button id="btnCreateAsset" class="btn primary">✨ Create Asset</button>
    </div>
  `;
  document.getElementById("btnCreateAsset")?.addEventListener("click", createAsset);
  if (overlay) overlay.classList.add("open");
});

async function createAsset() {
  const inputs = document.querySelectorAll("#drawerBody [data-field]");
  const payload = {};
  inputs.forEach(el => { payload[el.dataset.field] = el.value || null; });
  try {
    const a = await api("/api/assets", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload),
    });
    closeDrawer();
    refreshCurrentView();
    showToast(`Created asset: ${a.name || "Asset #" + a.id}`, false);
    openAssetDrawer(a.id);
  } catch (err) {
    showToast("Failed to create asset: " + err.message, true);
  }
}

// ---------------------------------------------------------------- Campaign Reconciliation
state.reconFilter = "all";
state.reconSearch = "";
state.reconCategory = "";

function renderVariancePill(days) {
  if (days === null || days === undefined) {
    return `<span class="variance-pill" style="background:rgba(255,255,255,0.05); color:var(--text-faint);">No Due Date</span>`;
  }
  if (Math.abs(days) <= 60) {
    return `<span class="variance-pill ok" title="Aligned with Master statutory date within 60 days">🟢 ALIGNED (${days > 0 ? '+' : ''}${days}d)</span>`;
  }
  if (days > 60) {
    return `<span class="variance-pill late" title="Planned later than Master statutory date by ${days} days">🔴 LATE +${days}d</span>`;
  }
  return `<span class="variance-pill ahead" title="Planned ahead of Master statutory date by ${Math.abs(days)} days">🟣 AHEAD ${days}d</span>`;
}

function renderProgressBadge(val) {
  const v = (val || "").trim();
  if (!v) return `<span class="badge-progress" style="background:rgba(255,255,255,0.05); color:var(--text-faint);">—</span>`;
  const lower = v.toLowerCase();
  if (lower === "done") {
    return `<span class="badge-progress done">✅ DONE</span>`;
  }
  if (lower === "holding") {
    return `<span class="badge-progress holding">⏳ HOLDING</span>`;
  }
  if (lower.includes("survice") || lower.includes("out of") || lower.includes("service") || lower.includes("oos")) {
    return `<span class="badge-progress oos">OUT OF SERVICE</span>`;
  }
  if (lower.includes("replaced")) {
    return `<span class="badge-progress oos">REPLACED</span>`;
  }
  // Junk text (data-entry artifacts like 'tubing material') must not look like
  // a real progress value — render it as neutral data, not a progress badge.
  const known = ["done", "holding", "survice", "out of", "service", "oos", "replaced"];
  if (!known.some(k => lower.includes(k))) {
    return `<span class="badge-progress" style="background:rgba(255,255,255,0.04); color:var(--text-faint); border:1px dashed rgba(255,255,255,0.15);" title="Unrecognized value — check source data">${esc(v)}</span>`;
  }
  return `<span class="badge-progress" style="background:rgba(0,210,255,0.12); color:var(--accent);">${esc(v)}</span>`;
}
const formatUtProgressBadge = renderProgressBadge;

window.setReconFilter = function(filterName) {
  state.reconFilter = filterName;
  document.querySelectorAll(".subnav-btn[data-recon-filter]").forEach(b => {
    b.classList.toggle("active", b.dataset.reconFilter === filterName);
  });
  loadReconciliationItems();
};

async function loadReconciliation() {
  try {
    const stats = await api("/api/reconciliation");
    
    // Update Header KPIs
    const matchEl = document.getElementById("kpiReconMatchRate");
    if (matchEl) matchEl.textContent = `${stats.match_rate}%`;
    const utEl = document.getElementById("kpiReconUtRate");
    if (utEl) utEl.textContent = `${stats.progress_stats?.ut_percent || 0}%`;

    // Update Cards
    const totEl = document.getElementById("reconTotalItems");
    if (totEl) totEl.textContent = stats.total_items;
    const matEl = document.getElementById("reconMatchedItems");
    if (matEl) matEl.textContent = `${stats.matched_count} (${stats.match_rate}%)`;
    const varEl = document.getElementById("reconVariances");
    if (varEl) varEl.textContent = `${stats.variance_count} (${stats.delayed_count} Late)`;
    const omiEl = document.getElementById("reconOmissions");
    if (omiEl) omiEl.textContent = stats.omitted_overdue_count;
    const unmEl = document.getElementById("reconUnmatched");
    if (unmEl) unmEl.textContent = stats.unmatched_count;

    // Update Subnav Count Badges
    const subAll = document.getElementById("countSubAll"); if (subAll) subAll.textContent = stats.total_items;
    const subVar = document.getElementById("countSubVariances"); if (subVar) subVar.textContent = stats.variance_count;
    const subOmi = document.getElementById("countSubOmissions"); if (subOmi) subOmi.textContent = stats.omitted_overdue_count;
    const subDone = document.getElementById("countSubDone"); if (subDone) subDone.textContent = stats.progress_stats?.ut_done || 0;
    const subHold = document.getElementById("countSubHolding"); if (subHold) subHold.textContent = stats.progress_stats?.ut_holding || 0;
    const subUnm = document.getElementById("countSubUnmatched"); if (subUnm) subUnm.textContent = stats.unmatched_count;

    // Update Bulk Sync Master Button
    const btnSync = document.getElementById("btnBulkSyncMaster");
    if (btnSync) {
      const pending = stats.pending_sync_count ?? 0;
      const synced = stats.synced_count ?? 0;
      if (pending > 0) {
        btnSync.innerHTML = `⚡ Sync Completed to Master (${pending} Pending)`;
        btnSync.disabled = false;
        btnSync.style.opacity = "1";
        btnSync.style.cursor = "pointer";
      } else if (synced > 0) {
        btnSync.innerHTML = `✅ Master Synchronized (${synced} Assets)`;
        btnSync.disabled = true;
        btnSync.style.opacity = "0.75";
        btnSync.style.cursor = "default";
      } else {
        btnSync.innerHTML = `⚡ Sync Completed to Master (0)`;
        btnSync.disabled = true;
        btnSync.style.opacity = "0.6";
        btnSync.style.cursor = "not-allowed";
      }
    }

    loadReconciliationItems();
  } catch (err) {
    showToast("Failed to load reconciliation stats: " + err.message, true);
  }
}

async function loadReconciliationItems() {
  const wrap = document.getElementById("reconciliationTableWrap");
  if (!wrap) return;

  try {
    const metaText = document.getElementById("reconMetaText");
    const q = document.getElementById("reconSearchBox")?.value?.trim() || "";
    const cat = document.getElementById("reconCategoryFilter")?.value || "";

    if (state.reconFilter === "omissions") {
      // Load Master Overdue Assets Omitted from Campaign
      const omissions = await api("/api/reconciliation/omissions" + (q ? `?q=${encodeURIComponent(q)}` : ""));
      if (metaText) metaText.textContent = `${omissions.length} overdue/high-risk master asset(s) not covered in campaign${q ? ` (matching "${q}")` : ""}`;
      wrap.innerHTML = renderOmissionsTable(omissions);
      bindRowClicks(wrap);
      return;
    }

    const params = new URLSearchParams({
      filter: state.reconFilter,
      q: q,
      category: cat,
    });

    const items = await api("/api/refined_plan?" + params.toString());
    
    // Populate categories in filter if empty
    const catSelect = document.getElementById("reconCategoryFilter");
    if (catSelect && catSelect.options.length <= 1 && items.length) {
      const cats = Array.from(new Set(items.map(it => it.category).filter(Boolean))).sort();
      cats.forEach(c => {
        const opt = document.createElement("option");
        opt.value = c;
        opt.textContent = c;
        catSelect.appendChild(opt);
      });
      catSelect.value = cat;
    }

    if (metaText) {
      metaText.textContent = `Showing ${items.length} item(s)${q ? ` matching "${q}"` : ""}${cat ? ` in ${cat}` : ""}`;
    }

    wrap.innerHTML = renderReconciliationTable(items);
    
    // Bind click to open Asset Drawer for matched master assets
    wrap.querySelectorAll("tr[data-asset-id]").forEach(tr => {
      const aid = tr.dataset.assetId;
      if (aid) {
        tr.addEventListener("click", () => openAssetDrawer(parseInt(aid)));
      }
    });

  } catch (err) {
    showToast("Failed to load campaign reconciliation: " + err.message, true);
  }
}

window.clearReconCategory = function() {
  const catSelect = document.getElementById("reconCategoryFilter");
  if (catSelect) catSelect.value = "";
  loadReconciliationItems();
};

function renderReconciliationTable(rows) {
  if (!rows || !rows.length) {
    const q = document.getElementById("reconSearchBox")?.value?.trim() || "";
    const cat = document.getElementById("reconCategoryFilter")?.value || "";
    const filter = state.reconFilter;
    let hint = "";
    if (q) {
      if (cat) {
        hint += ` <button class="btn ghost small" style="margin-left:8px;" onclick="clearReconCategory()">Clear Category Filter</button>`;
      }
      if (filter && filter !== "all") {
        hint += ` <button class="btn ghost small" style="margin-left:8px;" onclick="setReconFilter('all')">Search All Campaign Items</button>`;
      }
      return `<table class="assets"><tbody><tr class="empty-row"><td>No campaign items found matching "<strong>${esc(q)}</strong>".${hint}</td></tr></tbody></table>`;
    }
    return `<table class="assets"><tbody><tr class="empty-row"><td>No campaign items found matching this filter.</td></tr></tbody></table>`;
  }

  return `<table class="assets">
    <thead>
      <tr>
        <th>Campaign Item / Package</th>
        <th>Matched Master Asset</th>
        <th>Campaign Planned Date</th>
        <th>Master Statutory Due</th>
        <th>Variance</th>
        <th>Scope</th>
        <th>UT / Report Progress</th>
        <th>Scope & Recommendations</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(r => {
        const hasMatch = r.match_status === "MATCHED" && r.asset_id;
        const rowClass = hasMatch ? "recon-row-matched clickable-row" : "recon-row-unmatched";
        const isLate = r.date_variance_days && r.date_variance_days > 60;
        const isEarly = r.date_variance_days && r.date_variance_days < -60;
        
        let varBadge = `<span class="badge-variance on-time">ON SCHEDULE</span>`;
        if (isLate) {
          varBadge = `<span class="badge-variance delayed" title="Delayed by ${r.date_variance_days} days after statutory due date">⚠️ ${r.date_variance_days}d LATE</span>`;
        } else if (isEarly) {
          varBadge = `<span class="badge-variance proactive" title="Executing ${Math.abs(r.date_variance_days)} days ahead of statutory due date">🛡️ ${Math.abs(r.date_variance_days)}d EARLY</span>`;
        } else if (r.date_variance_days !== null && r.date_variance_days !== undefined) {
          varBadge = `<span class="badge-variance on-time">${r.date_variance_days >= 0 ? '+' : ''}${r.date_variance_days}d</span>`;
        }

        return `
        <tr class="${rowClass}" ${hasMatch ? `data-asset-id="${r.asset_id}"` : ''}>
          <td>
            <div style="font-weight:700; color:#fff;">${esc(r.extracted_tag || r.pkg_or_asset_no || '—')}</div>
            <div style="font-size:11.5px; color:var(--text-muted); line-height:1.35; margin-top:2px;">${esc(r.item_description || '—')}</div>
            <div style="margin-top:4px;">
              <span class="tag-mono" style="font-size:10px;">${esc(r.category || 'Package')}</span>
              ${r.pkg_or_asset_no && r.pkg_or_asset_no !== r.extracted_tag ? `<span class="tag-mono" style="font-size:10px; margin-left:4px;">Pkg: ${esc(r.pkg_or_asset_no)}</span>` : ''}
            </div>
          </td>
          <td>
            ${hasMatch ? `
              <div style="font-weight:600; color:var(--accent);">
                ${esc(r.master_name || 'Master Asset')}
              </div>
              <div class="tag-mono" style="color:#a5d6a7; margin-top:2px;">
                Tag: ${esc(r.master_tag || '—')} | Sheet: ${esc(r.master_source_sheet || '—')}
              </div>
              <div style="display:flex; gap:6px; margin-top:4px; align-items:center;">
                <span class="badge-pill-verified" title="Reconciled to master statutory asset">✓ MATCHED</span>
                ${r.master_risk ? `<span class="risk-badge risk-${r.master_risk.toLowerCase()}">${esc(r.master_risk)}</span>` : ''}
                ${r.master_location ? `<span class="tag-mono" style="font-size:10px; color:var(--text-muted);">Loc: ${esc(r.master_location)}</span>` : ''}
              </div>
            ` : `
              <div style="color:var(--text-muted); font-style:italic;">No baseline asset matched</div>
              <div style="margin-top:4px;">
                <span class="badge-pill-unmatched" title="Item present in campaign but not mapped to master statutory asset">UNMATCHED IN MASTER</span>
              </div>
            `}
          </td>
          <td>
            <div style="font-family:var(--font-mono); font-weight:600; color:#fff;">
              ${fmtDate(r.planned_insp_date)}
            </div>
            ${r.last_insp_date ? `<div style="font-size:10.5px; color:var(--text-muted);">Last: ${fmtDate(r.last_insp_date)}</div>` : ''}
          </td>
          <td>
            ${hasMatch ? `
              <div style="font-family:var(--font-mono); font-size:12px; font-weight:600; color:${r.master_internal_next ? '#bb86fc' : '#00d2ff'};">
                ${fmtDate(r.master_internal_next || r.master_osi_next)}
              </div>
              <div style="font-size:10.5px; color:var(--text-muted);">
                ${r.master_internal_next ? 'Statutory Internal' : 'Statutory OSI'}
              </div>
            ` : `<span style="color:var(--text-muted);">—</span>`}
          </td>
          <td>${hasMatch ? varBadge : '<span style="color:var(--text-muted);">—</span>'}</td>
          <td>
            <span class="badge-scope">${esc(r.scope_category || 'General')}</span>
            ${r.priority ? `<div style="font-size:10.5px; margin-top:3px; color:var(--text-muted);">Pri: ${esc(r.priority)}</div>` : ''}
          </td>
          <td>
            <div>${renderProgressBadge(r.ut_progress)}</div>
            ${r.report_issued && r.report_issued !== r.ut_progress ? `
              <div style="margin-top:4px; font-size:11px; color:var(--text-muted);">
                Report: ${renderProgressBadge(r.report_issued)}
              </div>
            ` : ''}
            ${r.synced_to_master ? `<div style="margin-top:4px;"><span class="badge-progress" style="background:rgba(0,230,118,0.18); color:#00e676; border:1px solid rgba(0,230,118,0.4); font-weight:700; font-size:10px;">✅ Master Synced</span></div>` : ""}
          </td>
          <td>
            <div style="font-size:11.5px; color:var(--text-muted); line-height:1.4; max-width:280px;">
              ${esc(r.insp_scope_remarks || r.remarks || r.master_remarks || '—')}
            </div>
          </td>
        </tr>`;
      }).join("")}
    </tbody>
  </table>`;
}

function renderOmissionsTable(rows) {
  if (!rows || !rows.length) {
    const q = document.getElementById("reconSearchBox")?.value?.trim() || "";
    if (q) {
      return `<table class="assets"><tbody><tr class="empty-row"><td>No omitted overdue assets found matching "<strong>${esc(q)}</strong>".</td></tr></tbody></table>`;
    }
    return `<table class="assets"><tbody><tr class="empty-row"><td>Zero omissions! All overdue master assets are accounted for in the campaign.</td></tr></tbody></table>`;
  }

  return `<table class="assets">
    <thead>
      <tr>
        <th>Omitted Master Asset</th>
        <th>System / Sheet</th>
        <th>Field / Location</th>
        <th>Statutory Due Date</th>
        <th>Compliance Status</th>
        <th>Remarks</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(a => `
        <tr data-id="${a.id}">
          <td>
            <div style="font-weight:600; color:#fff;">${esc(a.name || '—')} ${riskBadge(a)}</div>
            <div class="tag-mono">${esc(a.tag || a.sn || '—')}</div>
          </td>
          <td><span class="tag-mono" style="color:var(--accent); font-weight:600;">${esc(a.source_sheet || '—')}</span></td>
          <td><span class="tag-mono">${esc(a.plant || a.field || a.location || '—')}</span></td>
          <td>
            <div style="font-family:var(--font-mono); font-size:12px; color:var(--danger);">${fmtDate(a.next_due)}</div>
            ${a.date_internal_next ? `<div style="font-size:10.5px; color:#bb86fc;">Internal: ${fmtDate(a.date_internal_next)}</div>` : ''}
          </td>
          <td>${statusPill(a)}</td>
          <td><div style="max-width:260px; font-size:11.5px; color:var(--text-muted); overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${esc(a.remarks || '')}">${esc(a.remarks || '—')}</div></td>
        </tr>
      `).join("")}
    </tbody>
  </table>`;
}

// Reconciliation Toolbar Listeners
let reconSearchDebounce = null;
const reconSearchBox = document.getElementById("reconSearchBox");
if (reconSearchBox) {
  reconSearchBox.addEventListener("input", () => {
    clearTimeout(reconSearchDebounce);
    reconSearchDebounce = setTimeout(loadReconciliationItems, 160);
  });
  reconSearchBox.addEventListener("search", () => {
    clearTimeout(reconSearchDebounce);
    loadReconciliationItems();
  });
  reconSearchBox.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      clearTimeout(reconSearchDebounce);
      loadReconciliationItems();
    }
  });
}
document.getElementById("reconCategoryFilter")?.addEventListener("change", loadReconciliationItems);
document.getElementById("btnReloadRecon")?.addEventListener("click", loadReconciliation);
document.getElementById("btnExportRecon")?.addEventListener("click", () => {
  const cat = document.getElementById("reconCategoryFilter")?.value || "";
  const q = document.getElementById("reconSearchBox")?.value?.trim() || "";
  const params = new URLSearchParams({
    filter: state.reconFilter,
    category: cat,
    q: q,
  });
  window.location.href = "/api/reconciliation/export.csv?" + params.toString();
});

// Refined Plan Import Handler
const btnImportRefined = document.getElementById("btnImportRefined");
const refinedFileInput = document.getElementById("refinedFileInput");
if (btnImportRefined && refinedFileInput) {
  btnImportRefined.addEventListener("click", () => {
    refinedFileInput.value = "";
    refinedFileInput.click();
  });

  refinedFileInput.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const origText = btnImportRefined.innerHTML;
    btnImportRefined.disabled = true;
    btnImportRefined.innerHTML = "⏳ Importing…";

    try {
      const reader = new FileReader();
      reader.onload = async (ev) => {
        try {
          const base64Data = ev.target.result.split(",")[1];
          const res = await api("/api/reconciliation/import", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              filename: file.name,
              filedata: base64Data,
            }),
          });
          showToast(res.message || "Refined plan imported successfully!", false);
          loadReconciliation();
        } catch (err) {
          showToast("Import failed: " + err.message, true);
        } finally {
          btnImportRefined.disabled = false;
          btnImportRefined.innerHTML = origText;
        }
      };
      reader.onerror = () => {
        showToast("Failed to read file.", true);
        btnImportRefined.disabled = false;
        btnImportRefined.innerHTML = origText;
      };
      reader.readAsDataURL(file);
    } catch (err) {
      showToast("Error processing file: " + err.message, true);
      btnImportRefined.disabled = false;
      btnImportRefined.innerHTML = origText;
    }
  });
}

// Bulk Synchronize Completed Inspections to Master Plan
const btnBulkSyncMaster = document.getElementById("btnBulkSyncMaster");
if (btnBulkSyncMaster) {
  btnBulkSyncMaster.addEventListener("click", async () => {
    const confirmMsg = "⚡ Bulk Synchronize Completed Inspections to Master Plan\n\n" +
      "This action will:\n" +
      "1. Log official inspection records in the Master Plan for all campaign items marked as 'Done'.\n" +
      "2. Advance their next statutory due dates (OSI / Internal).\n" +
      "3. Officially clear their Overdue status across all dashboard metrics and reports.\n" +
      "4. Create an automatic backup of the database beforehand.\n\n" +
      "Do you wish to proceed?";

    if (!confirm(confirmMsg)) return;

    const origText = btnBulkSyncMaster.innerHTML;
    btnBulkSyncMaster.disabled = true;
    btnBulkSyncMaster.innerHTML = "⏳ Synchronizing…";

    try {
      const res = await api("/api/reconciliation/sync_completed", { method: "POST" });
      showToast(res.message || "Inspections synchronized successfully!", false);
      loadReconciliation();
      loadDashboard();
      if (state.view === "assets") loadAssets();
    } catch (err) {
      showToast("Sync failed: " + err.message, true);
    } finally {
      btnBulkSyncMaster.disabled = false;
      btnBulkSyncMaster.innerHTML = origText;
    }
  });
}

// Clear Reconciled Campaign Data Handler
const btnClearRecon = document.getElementById("btnClearRecon");
if (btnClearRecon) {
  btnClearRecon.addEventListener("click", async () => {
    const confirmMsg = "Are you sure you want to clear all reconciled campaign data?\n\n" +
      "This will remove the campaign spreadsheet data and restore the view to its initial empty state.\n" +
      "Your Master Plan assets and statutory records will NOT be deleted.\n\n" +
      "Clear reconciled campaign?";

    if (!confirm(confirmMsg)) return;

    const origText = btnClearRecon.innerHTML;
    btnClearRecon.disabled = true;
    btnClearRecon.innerHTML = "⏳ Clearing…";

    try {
      const res = await api("/api/reconciliation/clear", { method: "POST" });
      showToast(res.message || "Reconciled campaign data cleared cleanly.", false);
      loadReconciliation();
    } catch (err) {
      showToast("Failed to clear campaign: " + err.message, true);
    } finally {
      btnClearRecon.disabled = false;
      btnClearRecon.innerHTML = origText;
    }
  });
}

// ---------------------------------------------------------------- Initialization
loadDashboard();