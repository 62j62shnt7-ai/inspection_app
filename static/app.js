/**
 * Master Inspection Plan — Client Application Logic
 */

const state = {
  view: "dashboard",
  assets: [],
  filteredAssets: [],
  page: 1,
  pageSize: 50,
  groupBy: "",
  sheets: [],
  editingAsset: null,
  activeDrawerTab: "specs",
};

// ---------------------------------------------------------------- Toast Notifications
function showToast(msg, isError = true) {
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
  return `<tr data-id="${a.id}">
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
  const extraHeads = compact ? "" : `<th>Pack # / Header</th><th>Plant & Location</th>`;
  return `<table class="assets">
    <thead><tr><th>Asset & Risk</th><th>Source Sheet</th><th>Status</th>${extraHeads}</tr></thead>
    <tbody>${rows.map(a => assetRowHtml(a, compact)).join("")}</tbody>
  </table>`;
}

function bindRowClicks(container) {
  if (!container) return;
  container.querySelectorAll("tr[data-id]").forEach(tr => {
    tr.addEventListener("click", () => openAssetDrawer(parseInt(tr.dataset.id)));
  });
}

async function loadDashboard() {
  try {
    const d = await api("/api/dashboard");
    
    // 1. Stat Cards
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

    // 2. High Risk List
    const highRiskEl = document.getElementById("highRiskList");
    if (highRiskEl) {
      highRiskEl.innerHTML = tableHtml(d.high_risk || [], true);
      bindRowClicks(highRiskEl);
      const countEl = document.getElementById("highRiskCountBadge");
      if (countEl) countEl.textContent = (d.high_risk || []).length;
    }

    // 3. Overdue List
    const overdueEl = document.getElementById("overdueList");
    if (overdueEl) {
      overdueEl.innerHTML = tableHtml(d.overdue || [], true);
      bindRowClicks(overdueEl);
      const countEl = document.getElementById("overdueCountBadge");
      if (countEl) countEl.textContent = (d.overdue || []).length;
    }

    // 4. Temporary Repairs Widget List
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

    // 5. Critical Scope Widget List
    const criticalWidgetEl = document.getElementById("criticalWidgetList");
    if (criticalWidgetEl) {
      const crit = d.pending_critical || [];
      const countEl = document.getElementById("criticalCountBadge");
      if (countEl) countEl.textContent = crit.length;
      if (!crit.length) {
        criticalWidgetEl.innerHTML = `<table class="assets"><tbody><tr class="empty-row"><td>No pending critical turnaround replacements.</td></tr></tbody></table>`;
      } else {
        criticalWidgetEl.innerHTML = `<table class="assets">
          <thead><tr><th>Pack # & Item</th><th>Category Section</th><th>Replacement Scope</th><th>Status</th></tr></thead>
          <tbody>${crit.map(c => `<tr>
            <td><div style="font-weight:600; color:#fff;">${esc(c.item_description || c.pack_no || "")}</div><div class="tag-mono">${esc(c.pack_no || "")}</div></td>
            <td><span class="tag-mono">${esc(c.category_section || "")}</span></td>
            <td><div style="max-width:240px; font-size:12px; color:var(--text-muted);">${esc(c.replacement_scope || "")}</div></td>
            <td><span class="status-pill soon">Pending</span></td>
          </tr>`).join("")}</tbody>
        </table>`;
      }
    }

    // 6. Source Sheet Chips
    const bySheetEl = document.getElementById("bySheet");
    if (bySheetEl && d.sheets && d.by_sheet) {
      state.sheets = d.sheets;
      bySheetEl.innerHTML = d.sheets.map(s => {
        const info = d.by_sheet[s] || {total:0, overdue:0};
        return `<span class="chip" onclick="filterBySheet('${esc(s)}')">
          <b>${esc(s)}</b> — ${info.total} ${info.overdue ? `· <span class="od">${info.overdue} overdue</span>` : ""}
        </span>`;
      }).join("");

      // Update sheet select options
      const sel = document.getElementById("sheetFilter");
      if (sel) {
        const current = sel.value;
        sel.innerHTML = `<option value="">All Sheets (${d.sheets.length} Systems)</option>` + 
          d.sheets.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
        sel.value = current;
      }
    }

  } catch (err) {
    showToast("Failed to load dashboard: " + err.message);
  }
}

window.filterBySheet = function(sheetName) {
  const sel = document.getElementById("sheetFilter");
  if (sel) sel.value = sheetName;
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
document.getElementById("groupFilter")?.addEventListener("change", () => { state.page = 1; loadAssets(); });
document.getElementById("overdueOnly")?.addEventListener("change", () => { state.page = 1; loadAssets(); });

document.getElementById("btnResetFilters")?.addEventListener("click", () => {
  if (searchBox) searchBox.value = "";
  if (btnClearSearch) btnClearSearch.style.display = "none";
  const sf = document.getElementById("sheetFilter"); if (sf) sf.value = "";
  const rf = document.getElementById("riskFilter"); if (rf) rf.value = "";
  const gf = document.getElementById("groupFilter"); if (gf) gf.value = "";
  const oo = document.getElementById("overdueOnly"); if (oo) oo.checked = false;
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
    const overdue = document.getElementById("overdueOnly")?.checked ? "1" : "0";
    state.groupBy = document.getElementById("groupFilter")?.value || "";

    const params = new URLSearchParams({q, sheet, risk, overdue});
    const rows = await api("/api/assets?" + params.toString());
    state.assets = rows;
    state.filteredAssets = rows;

    renderAssetsTable();
  } catch (err) {
    showToast("Failed to load assets: " + err.message);
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

  // Handle grouping if selected
  if (state.groupBy) {
    const groups = {};
    pageRows.forEach(a => {
      const key = a[state.groupBy] || "Unassigned";
      if (!groups[key]) groups[key] = [];
      groups[key].push(a);
    });

    let html = `<table class="assets">
      <thead><tr><th>Asset & Risk</th><th>Source Sheet</th><th>Status</th><th>Pack # / Header</th><th>Plant & Location</th></tr></thead>
      <tbody>`;

    for (const [groupName, groupAssets] of Object.entries(groups)) {
      html += `<tr class="group-header-row"><td colspan="5">📂 ${esc(groupName)} (${groupAssets.length} assets)</td></tr>`;
      html += groupAssets.map(a => assetRowHtml(a, false)).join("");
    }
    html += `</tbody></table>`;
    wrapEl.innerHTML = html;
  } else {
    wrapEl.innerHTML = tableHtml(pageRows, false);
  }

  bindRowClicks(wrapEl);
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
    showToast("Failed to load critical turnaround assets: " + err.message);
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
    showToast("Failed to update status: " + err.message);
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
    showToast("Failed to load temporary repairs: " + err.message);
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
    showToast("Failed to load yearly plan: " + err.message);
  }
}

// ---------------------------------------------------------------- Export & Re-import
document.getElementById("btnExport")?.addEventListener("click", () => {
  const sheet = document.getElementById("sheetFilter")?.value || "";
  const params = new URLSearchParams(sheet ? {sheet} : {});
  window.location.href = "/api/export.csv?" + params.toString();
});

const picker = document.getElementById("excelFilePicker");
const btnReimport = document.getElementById("btnReimport");

if (btnReimport && picker) {
  btnReimport.addEventListener("click", () => {
    picker.value = "";
    picker.click();
  });

  picker.addEventListener("change", async () => {
    const file = picker.files[0];
    if (!file) return;

    const msg = `Are you sure you want to re-import "${file.name}"?\n\nThis will refresh the SQLite database with the inspection sheets from this Excel file.`;
    if (!confirm(msg)) return;

    const origText = btnReimport.innerHTML;
    btnReimport.disabled = true;
    btnReimport.innerHTML = `<span>⏳</span> Importing…`;

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
          showToast(res.message || "Re-import successful!", false);
          refreshCurrentView();
        } catch (err) {
          showToast("Re-import failed: " + err.message);
        } finally {
          btnReimport.disabled = false;
          btnReimport.innerHTML = origText;
        }
      };
      reader.onerror = () => {
        showToast("Failed to read selected file.");
        btnReimport.disabled = false;
        btnReimport.innerHTML = origText;
      };
      reader.readAsDataURL(file);
    } catch (err) {
      showToast("Error preparing file upload: " + err.message);
      btnReimport.disabled = false;
      btnReimport.innerHTML = origText;
    }
  });
}

// ---------------------------------------------------------------- Slide-Over Asset Drawer
const overlay = document.getElementById("drawerOverlay");
document.getElementById("drawerClose")?.addEventListener("click", closeDrawer);
overlay?.addEventListener("click", e => { if (e.target === overlay) closeDrawer(); });

function closeDrawer() {
  if (overlay) overlay.classList.remove("open");
  state.editingAsset = null;
}

document.addEventListener("keydown", e => {
  if (e.key === "Escape") closeDrawer();
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

const RISK_FIELDS = [
  ["risk_category", "Risk Level (HIGH / MEDIUM / LOW)"],
  ["damage_mechanisms", "Damage Mechanisms (e.g. CUI, Internal Thinning)"],
  ["corrosion_rate", "Corrosion Rate (mm/yr)"],
  ["remaining_life", "Estimated Remaining Life (years)"],
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
        <button class="drawer-tab" id="tabbtn-risk" onclick="switchDrawerTab('risk')">🛡️ Risk & Condition</button>
        <button class="drawer-tab" id="tabbtn-logs" onclick="switchDrawerTab('logs')">🔍 NDT & Inspection Logs (${(a.log || []).length})</button>
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

        <div class="section-title">Scheduled Inspection Due Dates</div>
        <div class="field-grid">
          ${DATE_FIELDS.map(([k,l]) => fieldInput(k,l,a[k],"text")).join("")}
        </div>

        <div class="field-grid wide">
          <div class="field">
            <label>Remarks & Integrity Recommendations</label>
            <textarea data-field="remarks">${esc(a.remarks)}</textarea>
          </div>
        </div>

        ${extraRows.length ? `
          <div class="section-title">Original Sheet Column Metadata</div>
          <div class="field-grid wide">
            ${extraRows.map(([k,v]) => `<div class="field"><label>${esc(k)}</label>
              <input value="${esc(v)}" readonly></div>`).join("")}
          </div>` : ""}
      </div>

      <!-- TAB 3: LOGS & NDT -->
      <div id="dtab-logs" class="dtab-content" style="display:none;">
        <div class="section-title">Record New NDT / Inspection Activity</div>
        <div class="field-grid">
          <div class="field"><label>Inspection Date (YYYY-MM-DD)</label><input id="logDate" type="date"></div>
          <div class="field"><label>Inspection Type</label>
            <select id="logType">
              <option value="OSI">OSI (On-Stream Inspection)</option>
              <option value="Internal">Internal Inspection</option>
              <option value="OSI-ADV">OSI Advanced NDT (PAUT/TOFD)</option>
              <option value="NDT UT Thickness">NDT UT Wall Thickness</option>
              <option value="Visual / CUI">Visual / CUI Surveillance</option>
              <option value="Other">Other Evaluation</option>
            </select>
          </div>
        </div>
        <div class="field-grid">
          <div class="field"><label>Inspector Name & Certification</label><input id="logInspector" placeholder="e.g. M. Ahmed (API 570 #4421)"></div>
          <div class="field"><label>NDT Technique</label><input id="logMethod" placeholder="e.g. Digital UT, MPT, Visual"></div>
        </div>
        <div class="field-grid">
          <div class="field"><label>Actual Measured Min Thickness t_act (mm)</label><input id="logTActual" placeholder="e.g. 5.8 mm"></div>
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
    showToast("Failed to open asset details: " + err.message);
  }
}

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
    showToast(`Asset ${archived ? "archived" : "restored"}`, false);
  } catch (err) {
    showToast("Archive toggle failed: " + err.message);
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
    showToast("Failed to add inspection log: " + err.message);
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
    showToast("Failed to delete log: " + err.message);
  }
};

function refreshCurrentView() {
  if (state.view === "dashboard") loadDashboard();
  if (state.view === "assets") loadAssets();
  if (state.view === "critical") loadCriticalAssets();
  if (state.view === "temprepairs") loadTempRepairs();
  if (state.view === "yearly") loadYearlyPlan();
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
    <div class="section-title">Risk Profile</div>
    <div class="field-grid">
      ${RISK_FIELDS.map(([k,l]) => fieldInput(k,l,"")).join("")}
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
    showToast("Failed to create asset: " + err.message);
  }
}

// ---------------------------------------------------------------- Initialization
loadDashboard();