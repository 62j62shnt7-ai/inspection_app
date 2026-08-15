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
  if (asset.is_deferred) {
    return `<span class="status-pill deferred" title="Approved MOC Deferral until ${asset.deferral_expiry || ''}">🛡️ DEFERRED (${fmtDate(asset.deferral_expiry)})</span>`;
  }
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
    loadYearlyPlan();
  }
}

// ---------------------------------------------------------------- Dashboard View
function assetRowHtml(a, compact) {
  const tag = a.tag || a.sn || "";
  const loc = [a.plant, a.location].filter(Boolean).join(" · ") || "—";
  const isSelected = state.selectedAssetIds.has(a.id);
  const checkTd = compact ? "" : `<td style="width:36px; text-align:center;" onclick="event.stopPropagation();">
    <input type="checkbox" class="row-checkbox" data-id="${a.id}" ${isSelected ? "checked" : ""}>
  </td>`;

  return `<tr data-id="${a.id}">
    ${checkTd}
    <td>
      <div style="font-weight:600; color:#fff;">${esc(a.name || "(unnamed)")} ${riskBadge(a)}</div>
      <div class="tag-mono">${esc(tag)}</div>
    </td>
    <td><span class="tag-mono">${esc(a.source_sheet || "")}</span></td>
    <td>${statusPill(a)}</td>
    ${compact ? "" : `<td><span class="tag-mono" style="font-weight:600; color:#fff;">${esc(a.field || "—")}</span></td><td>${esc(loc)}</td>`}
  </tr>`;
}

function tableHtml(rows, compact) {
  if (!rows || !rows.length) {
    return `<table class="assets"><tbody><tr class="empty-row"><td>No records found.</td></tr></tbody></table>`;
  }
  const checkTh = compact ? "" : `<th style="width:36px; text-align:center;"><input type="checkbox" id="selectAllCheckbox"></th>`;
  const extraHeads = compact ? "" : `<th>Pack # / Header</th><th>Plant & Location</th>`;
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
        <div class="stat-card overdue">
          <div class="n">${d.overdue_count}</div>
          <div class="label">Overdue Inspections</div>
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
      `;
    }

    // 4. Render 5x5 RBI Matrix
    if (d.rbi_matrix) {
      render5x5Matrix(d.rbi_matrix);
    }

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
    }

  } catch (err) {
    showToast("Failed to load dashboard: " + err.message, true);
  }
}

window.filterBySheet = function(sheetName) {
  const sel = document.getElementById("sheetFilter");
  if (sel) sel.value = sheetName;
  switchView("assets");
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
document.getElementById("agingFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("deferralFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("groupFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("overdueOnly")?.addEventListener("change", () => { state.page = 1; loadAssets(); });

document.getElementById("btnResetFilters")?.addEventListener("click", () => {
  if (searchBox) searchBox.value = "";
  if (btnClearSearch) btnClearSearch.style.display = "none";
  const sf = document.getElementById("sheetFilter"); if (sf) sf.value = "";
  const rf = document.getElementById("riskFilter"); if (rf) rf.value = "";
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
    const aging = document.getElementById("agingFilter")?.value || "";
    const deferral = document.getElementById("deferralFilter")?.value || "";
    const overdue = document.getElementById("overdueOnly")?.checked ? "1" : "0";
    state.groupBy = document.getElementById("groupFilter")?.value || "";

    const params = new URLSearchParams({
      q, sheet, risk, overdue,
      aging, deferral,
      rbi_cell: state.activeRbiCell
    });

    const rows = await api("/api/assets?" + params.toString());
    state.assets = rows;
    state.filteredAssets = rows;

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
document.getElementById("btnPrintYear")?.addEventListener("click", () => window.print());

async function loadYearlyPlan() {
  try {
    const yearInput = document.getElementById("yearInput");
    const year = yearInput ? yearInput.value : new Date().getFullYear();
    const rows = await api("/api/yearly_plan?year=" + encodeURIComponent(year));
    
    const metaEl = document.getElementById("yearlyMeta");
    if (metaEl) {
      metaEl.textContent = `${rows.length} asset(s) scheduled for inspection in ${year}`;
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

// ---------------------------------------------------------------- Initialization
loadDashboard();