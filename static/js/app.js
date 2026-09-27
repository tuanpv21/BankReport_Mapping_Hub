// BankReport Mapping Hub - Master-Detail Modern Frontend Logic

document.addEventListener("DOMContentLoaded", () => {
  let allReports = [];
  let currentSystemFilter = "ALL";
  let activeReport = null;
  let activeReportFields = [];
  let activeTab = "tab-dashboard";
  let currentParsedFields = [];

  const tabLinks = document.querySelectorAll(".nav-link");
  const tabViews = document.querySelectorAll(".tab-view");

  initNav();
  initGlobalSearch();
  initExplorer();
  initStudio();
  initLineage();
  initExportImport();
  initModals();
  initSync();

  loadStats();
  loadReports();

  function initNav() {
    tabLinks.forEach(link => {
      link.addEventListener("click", () => {
        const target = link.getAttribute("data-tab");
        switchTab(target);
      });
    });
  }

  function switchTab(tabId) {
    activeTab = tabId;
    tabLinks.forEach(l => l.classList.toggle("active", l.getAttribute("data-tab") === tabId));
    tabViews.forEach(v => v.classList.toggle("active", v.id === tabId));

    if (tabId === "tab-dashboard") loadStats();
    if (tabId === "tab-explorer" && !activeReport && allReports.length > 0) {
      selectReport(allReports[0]);
    }
    if (tabId === "tab-studio") {
      populateStudioDropdown();
    }
  }

  function showToast(msg, type = "info") {
    const hub = document.getElementById("toast-hub");
    const toast = document.createElement("div");
    toast.className = "toast-msg";
    let icon = "fa-circle-info text-blue";
    if (type === "success") icon = "fa-circle-check text-emerald";
    if (type === "error") icon = "fa-circle-xmark text-rose";
    toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${msg}</span>`;
    hub.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = "0";
      setTimeout(() => toast.remove(), 250);
    }, 3500);
  }

  function escapeHtml(text) {
    if (!text) return "";
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function truncate(str, max = 40) {
    if (!str) return "";
    return str.length > max ? str.substring(0, max) + "..." : str;
  }

  async function loadStats() {
    try {
      const res = await fetch("/api/stats");
      const data = await res.json();

      document.getElementById("stat-total-reports").textContent = data.total_reports || 0;
      document.getElementById("stat-reports-breakdown").textContent = `CIC: ${data.cic?.reports || 0} | TT35: ${data.tt35?.reports || 0}`;

      document.getElementById("stat-total-mappings").textContent = data.total_mappings || 0;
      document.getElementById("stat-mappings-breakdown").textContent = `CIC: ${data.cic?.mappings || 0} | TT35: ${data.tt35?.mappings || 0}`;

      const ratio = data.mapped_ratio || 0;
      document.getElementById("stat-mapped-ratio").textContent = `${ratio}%`;
      document.getElementById("stat-mapped-count").textContent = `${data.mapped_source || 0} chỉ tiêu đã map nguồn`;
      document.getElementById("bar-mapped-ratio").style.width = `${ratio}%`;

      document.getElementById("stat-unmapped-count").textContent = data.unmapped_source || 0;
      const unmappedRatio = Math.round((data.unmapped_source / (data.total_mappings || 1)) * 100);
      document.getElementById("bar-unmapped-ratio").style.width = `${unmappedRatio}%`;
    } catch (e) {
      console.error("Error loading stats:", e);
    }
  }

  async function loadReports() {
    try {
      const res = await fetch("/api/reports");
      allReports = await res.json();

      renderDashboardChips();
      renderMasterReportsList();
      populateExportDropdown();
      populateStudioDropdown();

      if (allReports.length > 0 && !activeReport) {
        selectReport(allReports[0]);
      }
    } catch (e) {
      console.error("Error loading reports:", e);
    }
  }

  function renderDashboardChips() {
    const cicBox = document.getElementById("dash-cic-chips");
    const tt35Box = document.getElementById("dash-tt35-chips");

    cicBox.innerHTML = "";
    tt35Box.innerHTML = "";

    allReports.forEach(r => {
      const chip = document.createElement("div");
      chip.className = "report-item-chip";
      chip.innerHTML = `
        <div class="chip-meta">
          <strong>${r.report_code}</strong>
          <span>${r.field_count} chỉ tiêu</span>
        </div>
        <p class="chip-desc">${escapeHtml(r.report_name || r.report_code)}</p>
      `;

      chip.addEventListener("click", () => {
        switchTab("tab-explorer");
        selectReport(r);
      });

      if (r.system_code === "CIC") {
        cicBox.appendChild(chip);
      } else {
        tt35Box.appendChild(chip);
      }
    });
  }

  function initExplorer() {
    document.querySelectorAll(".btn-toggle").forEach(btn => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".btn-toggle").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        currentSystemFilter = btn.getAttribute("data-sys");
        renderMasterReportsList();
      });
    });

    const filterInput = document.getElementById("report-filter-input");
    filterInput.addEventListener("input", () => {
      renderMasterReportsList();
    });

    const fieldInput = document.getElementById("field-filter-input");
    fieldInput.addEventListener("input", filterAndRenderFields);

    const statusSelect = document.getElementById("filter-field-status");
    statusSelect.addEventListener("change", filterAndRenderFields);

    document.getElementById("btn-export-active-report").addEventListener("click", () => {
      if (activeReport) {
        window.location.href = `/api/export/excel?system=${activeReport.system_code}&report=${activeReport.report_code}`;
      }
    });

    const btnJump = document.getElementById("btn-jump-to-studio");
    if (btnJump) {
      btnJump.addEventListener("click", () => {
        switchTab("tab-studio");
        if (activeReport) {
          const sel = document.getElementById("studio-select-report");
          if (sel) {
            sel.value = `${activeReport.system_code}::${activeReport.report_code}`;
            loadProcedureForSelectedReport();
          }
        }
      });
    }
  }

  function renderMasterReportsList() {
    const container = document.getElementById("master-reports-list");
    const searchVal = document.getElementById("report-filter-input").value.toLowerCase().trim();

    container.innerHTML = "";

    const filtered = allReports.filter(r => {
      const matchSys = currentSystemFilter === "ALL" || r.system_code === currentSystemFilter;
      const matchSearch = !searchVal || 
        r.report_code.toLowerCase().includes(searchVal) || 
        (r.report_name && r.report_name.toLowerCase().includes(searchVal));
      return matchSys && matchSearch;
    });

    if (filtered.length === 0) {
      container.innerHTML = '<div class="text-center py-4 text-muted" style="font-size:12px;">Không có mẫu biểu phù hợp</div>';
      return;
    }

    filtered.forEach(r => {
      const item = document.createElement("div");
      item.className = "master-report-card";
      if (activeReport && activeReport.report_code === r.report_code && activeReport.system_code === r.system_code) {
        item.classList.add("active");
      }

      const indClass = r.system_code === "CIC" ? "sys-ind-cic" : "sys-ind-tt35";

      item.innerHTML = `
        <div class="report-card-left">
          <div class="report-sys-indicator ${indClass}"></div>
          <div class="report-card-text">
            <h5>${r.report_code}</h5>
            <p title="${escapeHtml(r.report_name)}">${escapeHtml(r.report_name || r.report_code)}</p>
          </div>
        </div>
        <div class="report-card-badge">${r.field_count}</div>
      `;

      item.addEventListener("click", () => {
        document.querySelectorAll(".master-report-card").forEach(c => c.classList.remove("active"));
        item.classList.add("active");
        selectReport(r);
      });

      container.appendChild(item);
    });
  }

  async function selectReport(report) {
    activeReport = report;

    document.getElementById("active-report-badge").textContent = report.report_code;
    document.getElementById("active-report-title").textContent = `[${report.system_code}] ${report.report_code} - ${report.report_name || ''}`;
    document.getElementById("active-report-desc").textContent = `Bảng đích: ${report.target_table || report.report_code} | Quy định: ${report.regulation_ref || 'QĐ NHNN'}`;

    document.getElementById("field-filter-input").value = "";
    document.getElementById("filter-field-status").value = "ALL";

    const tbody = document.getElementById("fields-table-body");
    tbody.innerHTML = '<tr><td colspan="10" class="text-center py-5 text-muted"><i class="fa-solid fa-spinner fa-spin"></i> Đang tải dữ liệu chỉ tiêu...</td></tr>';

    try {
      const res = await fetch(`/api/mappings?system=${report.system_code}&report=${report.report_code}`);
      activeReportFields = await res.json();
      filterAndRenderFields();
    } catch (e) {
      tbody.innerHTML = '<tr><td colspan="10" class="text-center py-4 text-rose">Lỗi khi tải chỉ tiêu của báo cáo này.</td></tr>';
    }
  }

  // Render 9 standard columns: STT | Table | Column | DataType | Tên nghiệp vụ | Ghi chú điều kiện | Bảng nguồn | Cột nguồn | Công thức tính | Thao tác
  function filterAndRenderFields() {
    const tbody = document.getElementById("fields-table-body");
    const searchVal = document.getElementById("field-filter-input").value.toLowerCase().trim();
    const statusVal = document.getElementById("filter-field-status").value;

    const filtered = activeReportFields.filter(f => {
      const matchStatus = statusVal === "ALL" || f.status === statusVal;
      const matchSearch = !searchVal ||
        (f.field_code && f.field_code.toLowerCase().includes(searchVal)) ||
        (f.field_name_vi && f.field_name_vi.toLowerCase().includes(searchVal)) ||
        (f.source_table && f.source_table.toLowerCase().includes(searchVal)) ||
        (f.source_column && f.source_column.toLowerCase().includes(searchVal)) ||
        (f.transformation_rule && f.transformation_rule.toLowerCase().includes(searchVal)) ||
        (f.notes && f.notes.toLowerCase().includes(searchVal));
      return matchStatus && matchSearch;
    });

    document.getElementById("active-fields-count").textContent = `Hiển thị: ${filtered.length} / ${activeReportFields.length} chỉ tiêu`;

    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" class="text-center py-5 text-muted">Không tìm thấy chỉ tiêu nào phù hợp với từ khóa tìm kiếm.</td></tr>';
      return;
    }

    tbody.innerHTML = "";
    filtered.forEach((f, idx) => {
      const tr = document.createElement("tr");

      const tblName = f.target_table || (activeReport ? (activeReport.target_table || activeReport.report_code) : "");
      const colName = f.target_column || f.field_code;
      const dataType = f.data_type || "VARCHAR2(50)";
      const businessDesc = f.field_name_vi || f.field_code;
      const notes = f.notes || "-";

      let sourceTableHtml = "";
      if (f.source_table) {
        sourceTableHtml = `<span class="src-table-tag"><i class="fa-solid fa-database"></i> ${escapeHtml(f.source_table)}</span>`;
      } else {
        sourceTableHtml = '<span class="unmapped-tag"><i class="fa-solid fa-triangle-exclamation"></i> Chưa map</span>';
      }

      const sourceColHtml = f.source_column ? `<span class="src-col-tag">${escapeHtml(f.source_column)}</span>` : '<span class="text-subtle">-</span>';

      let ruleHtml = '<span class="text-subtle">-</span>';
      if (f.transformation_rule) {
        ruleHtml = `<span class="rule-code-snippet" title="${escapeHtml(f.transformation_rule)}">${escapeHtml(truncate(f.transformation_rule, 35))}</span>`;
      }

      tr.innerHTML = `
        <td class="text-center font-mono text-subtle" style="font-size:12px;">${idx + 1}</td>
        <td><span class="font-mono text-subtle" style="font-size:12px;">${escapeHtml(tblName)}</span></td>
        <td><span class="field-code-badge">${escapeHtml(colName)}</span></td>
        <td><span class="font-mono text-subtle" style="font-size:12px;">${escapeHtml(dataType)}</span></td>
        <td>
          <div class="field-name-title">${escapeHtml(businessDesc)}</div>
          ${f.lookup_ref ? `<div class="field-name-sub"><span class="tag-lookup"><i class="fa-solid fa-diagram-next"></i> ${escapeHtml(f.lookup_ref)}</span></div>` : ''}
        </td>
        <td><span class="text-subtle" style="font-size:12px;">${escapeHtml(notes)}</span></td>
        <td>${sourceTableHtml}</td>
        <td>${sourceColHtml}</td>
        <td>${ruleHtml}</td>
        <td class="text-right">
          <div class="action-btn-group" style="justify-content: flex-end;">
            <button class="btn-icon-sm btn-view-f" title="Xem chi tiết"><i class="fa-solid fa-eye"></i></button>
            <button class="btn-icon-sm btn-edit-f" title="Sửa chỉ tiêu"><i class="fa-solid fa-pen"></i></button>
          </div>
        </td>
      `;

      tr.querySelector(".btn-view-f").addEventListener("click", () => openDetailModal(f));
      tr.querySelector(".btn-edit-f").addEventListener("click", () => openEditModal(f));

      tbody.appendChild(tr);
    });
  }

  // ==========================================================
  // PROCEDURE & PACKAGE STUDIO LOGIC
  // ==========================================================
  function initStudio() {
    const reportSelector = document.getElementById("studio-select-report");
    const procNameInput = document.getElementById("studio-proc-name");
    const sqlTextarea = document.getElementById("studio-sql-input");
    const btnParse = document.getElementById("btn-parse-procedure");
    const btnLoadSample = document.getElementById("btn-load-sample-sql");
    const btnSaveCode = document.getElementById("btn-save-procedure-code");
    const btnApply = document.getElementById("btn-apply-parsed-mappings");

    if (reportSelector) {
      reportSelector.addEventListener("change", () => {
        loadProcedureForSelectedReport();
      });
    }

    if (btnLoadSample) {
      btnLoadSample.addEventListener("click", () => {
        const sampleSql = `-- Ví dụ Procedure nạp dữ liệu Khế ước vay CIC
insert into CIC_KU
  select a.customer_id TTC03, ---- Mã khách hàng
         a.account_number KU001, ---- Số khế ước vay
         to_char(a.value_date, 'YYYYMMDD') KU002, ---- Ngày giải ngân
         to_char(a.maturity_date, 'YYYYMMDD') KU003, ---- Ngày đáo hạn
         a.currency KU004, ---- Loại tiền vay
         a.credit_limit KU005, ---- Hạn mức tín dụng
         a.outstanding_balance KU006, ---- Dư nợ thực tế
         b.interest_rate KU007, ---- Lãi suất cho vay
         nvl(a.loan_purpose, '01') KU008 ---- Mục đích vay vốn
  from ODS_OD_ACCOUNT a
  left join ODS_INTEREST_RATE b on a.account_number = b.account_number
  where a.account_status = 'A';`;

        sqlTextarea.value = sampleSql;
        if (!procNameInput.value) procNameInput.value = "PKG_CIC_EXPORT.PRC_CIC_KU";
        showToast("Đã nạp mẫu cú pháp PL/SQL!", "info");
      });
    }

    if (btnSaveCode) {
      btnSaveCode.addEventListener("click", async () => {
        const selVal = reportSelector.value;
        if (!selVal) {
          alert("Vui lòng chọn một báo cáo trước khi lưu!");
          return;
        }
        const [sysCode, rptCode] = selVal.split("::");
        const procName = procNameInput.value.trim();
        const sqlText = sqlTextarea.value.trim();

        btnSaveCode.disabled = true;
        btnSaveCode.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang lưu...';

        try {
          const res = await fetch(`/api/reports/${rptCode}/procedure`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              system_code: sysCode,
              procedure_name: procName,
              procedure_code: sqlText
            })
          });
          const data = await res.json();
          if (data.success) {
            showToast(`Đã lưu Procedure cho báo cáo ${rptCode}!`, "success");
          } else {
            showToast("Lỗi khi lưu mã: " + (data.error || ""), "error");
          }
        } catch (e) {
          showToast("Lỗi kết nối máy chủ", "error");
        } finally {
          btnSaveCode.disabled = false;
          btnSaveCode.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> <span>Lưu mã Procedure</span>';
        }
      });
    }

    if (btnParse) {
      btnParse.addEventListener("click", async () => {
        const sqlText = sqlTextarea.value.trim();
        if (!sqlText) {
          alert("Vui lòng dán mã nguồn SQL trước khi bấm Bóc tách!");
          return;
        }

        const selVal = reportSelector.value;
        let sysCode = "CIC";
        let rptCode = "";
        if (selVal) {
          const parts = selVal.split("::");
          sysCode = parts[0];
          rptCode = parts[1];
        }

        btnParse.disabled = true;
        btnParse.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Đang bóc tách...</span>';

        try {
          const res = await fetch("/api/procedure/parse", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              sql: sqlText,
              system_code: sysCode,
              target_table: rptCode ? (sysCode === "CIC" ? `CIC_${rptCode}` : `RPTB_${rptCode}`) : ""
            })
          });

          const data = await res.json();
          if (data.success && data.fields) {
            currentParsedFields = data.fields;
            renderStudioPreview(data.fields);
            btnApply.disabled = (data.fields.length === 0);
            document.getElementById("studio-parsed-count").textContent = `${data.fields.length} chỉ tiêu`;
            showToast(`Bóc tách thành công ${data.fields.length} chỉ tiêu từ SQL!`, "success");
          } else {
            alert("Lỗi bóc tách: " + (data.error || "Không tìm thấy cấu trúc cột"));
          }
        } catch (e) {
          showToast("Lỗi máy chủ khi bóc tách cú pháp", "error");
        } finally {
          btnParse.disabled = false;
          btnParse.innerHTML = '<i class="fa-solid fa-bolt"></i> <span>Bóc tách Cú pháp SQL</span>';
        }
      });
    }

    if (btnApply) {
      btnApply.addEventListener("click", async () => {
        const selVal = reportSelector.value;
        if (!selVal) {
          alert("Vui lòng chọn báo cáo đích để áp dụng vào hệ thống!");
          return;
        }

        const [sysCode, rptCode] = selVal.split("::");
        const rows = readStudioPreviewRows();

        if (rows.length === 0) {
          alert("Không có chỉ tiêu nào trong danh sách xem trước để áp dụng!");
          return;
        }

        if (!confirm(`Bạn có chắc muốn áp dụng ${rows.length} chỉ tiêu vào báo cáo [${sysCode}] ${rptCode}?`)) {
          return;
        }

        btnApply.disabled = true;
        btnApply.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Đang áp dụng...</span>';

        try {
          const res = await fetch("/api/procedure/apply", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              system_code: sysCode,
              report_code: rptCode,
              target_table: rows[0]?.table || (sysCode === "CIC" ? `CIC_${rptCode}` : `RPTB_${rptCode}`),
              procedure_name: procNameInput.value.trim(),
              sql_text: sqlTextarea.value.trim(),
              fields: rows
            })
          });

          const data = await res.json();
          if (data.success) {
            showToast(`Đã cập nhật ${data.applied_count} chỉ tiêu vào cơ sở dữ liệu!`, "success");
            loadStats();
            loadReports();
            if (activeReport && activeReport.report_code === rptCode && activeReport.system_code === sysCode) {
              selectReport(activeReport);
            }
          } else {
            showToast("Lỗi khi áp dụng: " + (data.error || "Unknown"), "error");
          }
        } catch (e) {
          showToast("Lỗi kết nối máy chủ", "error");
        } finally {
          btnApply.disabled = false;
          btnApply.innerHTML = '<i class="fa-solid fa-circle-check"></i> <span>Áp dụng vào Bảng Mapping</span>';
        }
      });
    }
  }

  function populateStudioDropdown() {
    const sel = document.getElementById("studio-select-report");
    if (!sel || allReports.length === 0) return;

    const currentVal = sel.value;
    sel.innerHTML = "";
    allReports.forEach(r => {
      const opt = document.createElement("option");
      opt.value = `${r.system_code}::${r.report_code}`;
      opt.textContent = `[${r.system_code}] ${r.report_code} - ${r.report_name} (${r.field_count} cột)`;
      sel.appendChild(opt);
    });

    if (currentVal && Array.from(sel.options).some(o => o.value === currentVal)) {
      sel.value = currentVal;
    } else if (activeReport) {
      sel.value = `${activeReport.system_code}::${activeReport.report_code}`;
      loadProcedureForSelectedReport();
    }
  }

  async function loadProcedureForSelectedReport() {
    const sel = document.getElementById("studio-select-report");
    if (!sel || !sel.value) return;

    const [sysCode, rptCode] = sel.value.split("::");
    try {
      const res = await fetch(`/api/reports/${rptCode}/procedure?system=${sysCode}`);
      const data = await res.json();
      const nameInput = document.getElementById("studio-proc-name");
      const sqlInput = document.getElementById("studio-sql-input");

      if (nameInput) nameInput.value = data.procedure_name || "";
      if (sqlInput) sqlInput.value = data.procedure_code || "";
    } catch (e) {
      console.error("Error loading procedure code:", e);
    }
  }

  function renderStudioPreview(fields) {
    const tbody = document.getElementById("studio-preview-tbody");
    tbody.innerHTML = "";

    if (!fields || fields.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" class="text-center py-5 text-muted">Không bóc tách được trường nào từ câu lệnh SQL trên.</td></tr>';
      return;
    }

    fields.forEach((f, idx) => {
      const tr = document.createElement("tr");
      tr.className = "studio-preview-row";
      tr.innerHTML = `
        <td class="text-center font-mono text-subtle">${idx + 1}</td>
        <td><input type="text" class="input-clean input-inline-cell font-mono cell-table" value="${escapeHtml(f.table || '')}" placeholder="Table..."></td>
        <td><input type="text" class="input-clean input-inline-cell font-mono cell-column font-bold" value="${escapeHtml(f.column || '')}" placeholder="Column..."></td>
        <td><input type="text" class="input-clean input-inline-cell font-mono cell-datatype" value="${escapeHtml(f.data_type || 'VARCHAR2(50)')}"></td>
        <td><input type="text" class="input-clean input-inline-cell cell-name" value="${escapeHtml(f.field_name_vi || '')}" placeholder="Tên nghiệp vụ (Mô tả)..."></td>
        <td><input type="text" class="input-clean input-inline-cell cell-notes" value="${escapeHtml(f.notes || '')}" placeholder="Ghi chú điều kiện..."></td>
        <td><input type="text" class="input-clean input-inline-cell font-mono cell-srctable" value="${escapeHtml(f.source_table || '')}" placeholder="Bảng nguồn ODS..."></td>
        <td><input type="text" class="input-clean input-inline-cell font-mono cell-srccol" value="${escapeHtml(f.source_column || '')}" placeholder="Cột nguồn..."></td>
        <td><input type="text" class="input-clean input-inline-cell font-mono cell-rule" value="${escapeHtml(f.transformation_rule || '')}" placeholder="Công thức tính..."></td>
        <td class="text-center">
          <button class="btn-icon-sm text-rose btn-delete-row" title="Xóa dòng này"><i class="fa-solid fa-trash"></i></button>
        </td>
      `;

      tr.querySelector(".btn-delete-row").addEventListener("click", () => {
        tr.remove();
        const remaining = document.querySelectorAll(".studio-preview-row").length;
        document.getElementById("studio-parsed-count").textContent = `${remaining} chỉ tiêu`;
        if (remaining === 0) {
          document.getElementById("btn-apply-parsed-mappings").disabled = true;
          tbody.innerHTML = '<tr><td colspan="10" class="text-center py-4 text-muted">Đã xóa hết các dòng xem trước.</td></tr>';
        }
      });

      tbody.appendChild(tr);
    });
  }

  function readStudioPreviewRows() {
    const rows = [];
    document.querySelectorAll(".studio-preview-row").forEach(tr => {
      const table = tr.querySelector(".cell-table")?.value.trim() || "";
      const column = tr.querySelector(".cell-column")?.value.trim().toUpperCase() || "";
      const data_type = tr.querySelector(".cell-datatype")?.value.trim() || "VARCHAR2(50)";
      const field_name_vi = tr.querySelector(".cell-name")?.value.trim() || column;
      const notes = tr.querySelector(".cell-notes")?.value.trim() || "";
      const source_table = tr.querySelector(".cell-srctable")?.value.trim() || "";
      const source_column = tr.querySelector(".cell-srccol")?.value.trim() || "";
      const transformation_rule = tr.querySelector(".cell-rule")?.value.trim() || "";

      if (column) {
        rows.push({
          table,
          column,
          data_type,
          field_name_vi,
          notes,
          source_table,
          source_column,
          transformation_rule
        });
      }
    });
    return rows;
  }

  // ==========================================================
  // LINEAGE & IMPACT ANALYSIS
  // ==========================================================
  function initLineage() {
    const input = document.getElementById("lineage-search-field");
    const btn = document.getElementById("btn-do-lineage");

    const searchAction = async (val) => {
      const term = (val || input.value).trim();
      if (!term) return;
      input.value = term;

      const tbody = document.getElementById("lineage-results-body");
      tbody.innerHTML = '<tr><td colspan="7" class="text-center py-5 text-muted"><i class="fa-solid fa-spinner fa-spin"></i> Đang truy vết và phân tích tác động...</td></tr>';

      try {
        const res = await fetch(`/api/lineage?source_table=${encodeURIComponent(term)}`);
        const rows = await res.json();

        document.getElementById("lineage-count-tag").textContent = `${rows.length} chỉ tiêu ảnh hưởng`;
        document.getElementById("lineage-results-title").textContent = `Kết quả Phân tích Tác động cho "${term}"`;
        document.getElementById("lineage-summary-text").textContent = `Tìm thấy ${rows.length} chỉ tiêu báo cáo phụ thuộc vào bảng/cột này.`;

        if (rows.length === 0) {
          tbody.innerHTML = '<tr><td colspan="7" class="text-center py-5 text-muted">Không tìm thấy báo cáo hoặc chỉ tiêu nào sử dụng bảng này.</td></tr>';
          return;
        }

        tbody.innerHTML = "";
        rows.forEach(r => {
          const tr = document.createElement("tr");
          const sysBadge = r.system_code === "CIC" ? '<span class="status-badge" style="background:#e0f2fe;color:#0369a1;">CIC</span>' : '<span class="status-badge" style="background:#d1fae5;color:#065f46;">TT35</span>';

          tr.innerHTML = `
            <td><strong class="font-mono text-main">${escapeHtml(r.source_table)}</strong></td>
            <td><span class="font-mono text-subtle">${escapeHtml(r.source_column || '-')}</span></td>
            <td>${sysBadge}</td>
            <td><strong>${escapeHtml(r.report_code)}</strong></td>
            <td><span class="field-code-badge">${escapeHtml(r.field_code)}</span></td>
            <td><strong style="color:var(--text-main);">${escapeHtml(r.field_name_vi || '-')}</strong></td>
            <td><span class="font-mono text-subtle" style="font-size:11.5px;">${escapeHtml(r.implemented_by || '-')}</span></td>
          `;
          tbody.appendChild(tr);
        });
      } catch (e) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center py-4 text-rose">Lỗi trong quá trình phân tích tác động.</td></tr>';
      }
    };

    btn.addEventListener("click", () => searchAction());
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") searchAction();
    });

    document.querySelectorAll(".chip-suggest").forEach(c => {
      c.addEventListener("click", () => searchAction(c.getAttribute("data-term")));
    });
  }

  function initGlobalSearch() {
    const searchInput = document.getElementById("global-search-input");
    searchInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        const val = searchInput.value.trim();
        if (val) {
          switchTab("tab-explorer");
          document.getElementById("field-filter-input").value = val;
          filterAndRenderFields();
        }
      }
    });
  }

  function initExportImport() {
    document.getElementById("btn-header-export-excel").addEventListener("click", () => {
      window.location.href = "/api/export/excel?system=ALL";
    });
    document.getElementById("btn-header-export-word").addEventListener("click", () => {
      window.location.href = "/api/export/word?system=ALL";
    });

    document.getElementById("btn-export-excel-file").addEventListener("click", () => {
      const sys = document.getElementById("export-select-system").value;
      const rpt = document.getElementById("export-select-report").value;
      window.location.href = `/api/export/excel?system=${sys}&report=${rpt}`;
    });

    document.getElementById("btn-export-word-file").addEventListener("click", () => {
      const sys = document.getElementById("export-select-system").value;
      const rpt = document.getElementById("export-select-report").value;
      window.location.href = `/api/export/word?system=${sys}&report=${rpt}`;
    });

    const dropzone = document.getElementById("excel-dropzone");
    const fileInput = document.getElementById("excel-file-selector");
    const feedback = document.getElementById("import-upload-result");

    dropzone.addEventListener("click", () => fileInput.click());
    dropzone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropzone.style.borderColor = "var(--primary-600)";
    });
    dropzone.addEventListener("dragleave", () => {
      dropzone.style.borderColor = "var(--border-clean)";
    });
    dropzone.addEventListener("drop", (e) => {
      e.preventDefault();
      dropzone.style.borderColor = "var(--border-clean)";
      if (e.dataTransfer.files.length > 0) handleUpload(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener("change", () => {
      if (fileInput.files.length > 0) handleUpload(fileInput.files[0]);
    });

    async function handleUpload(file) {
      const formData = new FormData();
      formData.append("file", file);
      feedback.innerHTML = `<span class="text-primary"><i class="fa-solid fa-spinner fa-spin"></i> Đang tải lên và xử lý ${file.name}...</span>`;

      try {
        const res = await fetch("/api/import/excel", { method: "POST", body: formData });
        const data = await res.json();
        if (data.success) {
          feedback.innerHTML = `<span class="text-emerald"><i class="fa-solid fa-circle-check"></i> Đã nạp thành công ${data.imported_count} chỉ tiêu mới!</span>`;
          showToast(`Nhập thành công ${data.imported_count} chỉ tiêu`, "success");
          loadStats();
          loadReports();
        } else {
          feedback.innerHTML = `<span class="text-rose"><i class="fa-solid fa-circle-xmark"></i> Lỗi: ${data.error}</span>`;
        }
      } catch (e) {
        feedback.innerHTML = '<span class="text-rose">Lỗi kết nối máy chủ.</span>';
      }
    }
  }

  function populateExportDropdown() {
    const sel = document.getElementById("export-select-report");
    sel.innerHTML = '<option value="ALL">-- Tất cả mẫu biểu --</option>';
    allReports.forEach(r => {
      const opt = document.createElement("option");
      opt.value = r.report_code;
      opt.textContent = `[${r.system_code}] ${r.report_code} - ${r.report_name} (${r.field_count})`;
      sel.appendChild(opt);
    });
  }

  function initModals() {
    document.querySelectorAll(".btn-close-modal").forEach(btn => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".modal-backdrop").forEach(m => m.classList.remove("active"));
      });
    });

    document.getElementById("btn-open-create").addEventListener("click", () => {
      openCreateModal();
    });

    document.getElementById("btn-save-field").addEventListener("click", saveFieldForm);
  }

  function openDetailModal(f) {
    const modal = document.getElementById("modal-field-detail");
    document.getElementById("modal-view-sys-badge").textContent = f.system_code;
    document.getElementById("modal-view-title").textContent = `Chỉ tiêu [${f.report_code}] ${f.field_code}`;

    const body = document.getElementById("modal-view-body");
    body.innerHTML = `
      <div class="detail-item-grid">
        <div class="detail-label-txt">Mã & Tên Chỉ tiêu:</div>
        <div class="detail-val-txt">
          <span class="field-code-badge">${escapeHtml(f.field_code)}</span> 
          <strong style="font-size:14px; margin-left:8px;">${escapeHtml(f.field_name_vi || '-')}</strong>
        </div>

        <div class="detail-label-txt">Báo cáo & Bảng đích:</div>
        <div class="detail-val-txt">
          Mẫu <strong>${escapeHtml(f.report_code)}</strong> | Bảng: <span class="font-mono text-primary">${escapeHtml(f.target_table || '-')}</span>
        </div>

        <div class="detail-label-txt">Kiểu dữ liệu:</div>
        <div class="detail-val-txt font-mono">${escapeHtml(f.data_type || 'VARCHAR2(50)')}</div>

        <div class="detail-label-txt">Nguồn Dữ liệu:</div>
        <div class="detail-val-txt">
          ${f.source_table ? `<strong class="font-mono text-primary">${escapeHtml(f.source_table)}</strong> . <span class="font-mono">${escapeHtml(f.source_column || '-')}</span>` : '<span class="unmapped-tag"><i class="fa-solid fa-triangle-exclamation"></i> Chưa map nguồn</span>'}
        </div>

        <div class="detail-label-txt">Quy tắc / Công thức:</div>
        <div class="detail-val-txt">
          <div class="code-box-light">${escapeHtml(f.transformation_rule || '(Gán trực tiếp hoặc chưa ghi nhận quy tắc)')}</div>
        </div>

        <div class="detail-label-txt">Bảng mã quy đổi:</div>
        <div class="detail-val-txt">${escapeHtml(f.lookup_ref || 'Không sử dụng bảng mã')}</div>

        <div class="detail-label-txt">Thủ tục thực thi:</div>
        <div class="detail-val-txt font-mono">${escapeHtml(f.implemented_by || 'Chưa ghi nhận')}</div>

        <div class="detail-label-txt">Căn cứ quy định:</div>
        <div class="detail-val-txt">${escapeHtml(f.regulatory_ref || '-')}</div>

        <div class="detail-label-txt">Ghi chú điều kiện:</div>
        <div class="detail-val-txt">${escapeHtml(f.notes || '-')}</div>
      </div>
    `;

    document.getElementById("btn-trigger-edit-from-view").onclick = () => {
      modal.classList.remove("active");
      openEditModal(f);
    };

    modal.classList.add("active");
  }

  function openEditModal(f) {
    const modal = document.getElementById("modal-field-edit");
    document.getElementById("modal-form-title").textContent = `Chỉnh sửa Chỉ tiêu [${f.report_code}] ${f.field_code}`;

    document.getElementById("form-id").value = f.id;
    document.getElementById("form-system").value = f.system_code;
    document.getElementById("form-report").value = f.report_code;
    document.getElementById("form-field-code").value = f.field_code;
    document.getElementById("form-field-name-vi").value = f.field_name_vi || "";
    document.getElementById("form-data-type").value = f.data_type || "VARCHAR2(50)";
    document.getElementById("form-target-table").value = f.target_table || "";
    document.getElementById("form-status").value = f.status || "Active";
    document.getElementById("form-source-table").value = f.source_table || "";
    document.getElementById("form-source-column").value = f.source_column || "";
    document.getElementById("form-transformation-rule").value = f.transformation_rule || "";
    document.getElementById("form-lookup-ref").value = f.lookup_ref || "";
    document.getElementById("form-implemented-by").value = f.implemented_by || "";
    document.getElementById("form-regulatory-ref").value = f.regulatory_ref || "";
    document.getElementById("form-notes").value = f.notes || "";

    modal.classList.add("active");
  }

  function openCreateModal() {
    const modal = document.getElementById("modal-field-edit");
    document.getElementById("modal-form-title").textContent = "Thêm mới Chỉ tiêu Mapping";
    document.getElementById("form-field-mapping").reset();
    document.getElementById("form-id").value = "";

    if (activeReport) {
      document.getElementById("form-system").value = activeReport.system_code;
      document.getElementById("form-report").value = activeReport.report_code;
      document.getElementById("form-target-table").value = activeReport.target_table || "";
    }

    modal.classList.add("active");
  }

  async function saveFieldForm() {
    const id = document.getElementById("form-id").value;
    const payload = {
      system_code: document.getElementById("form-system").value,
      report_code: document.getElementById("form-report").value,
      field_code: document.getElementById("form-field-code").value,
      field_name_vi: document.getElementById("form-field-name-vi").value,
      data_type: document.getElementById("form-data-type").value,
      target_table: document.getElementById("form-target-table").value,
      status: document.getElementById("form-status").value,
      source_table: document.getElementById("form-source-table").value,
      source_column: document.getElementById("form-source-column").value,
      transformation_rule: document.getElementById("form-transformation-rule").value,
      lookup_ref: document.getElementById("form-lookup-ref").value,
      implemented_by: document.getElementById("form-implemented-by").value,
      regulatory_ref: document.getElementById("form-regulatory-ref").value,
      notes: document.getElementById("form-notes").value,
    };

    if (!payload.report_code || !payload.field_code || !payload.field_name_vi) {
      alert("Vui lòng điền đầy đủ Mã báo cáo, Mã chỉ tiêu và Tên nghiệp vụ!");
      return;
    }

    try {
      const url = id ? `/api/mappings/${id}` : "/api/mappings";
      const method = id ? "PUT" : "POST";
      const res = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.success) {
        showToast("Đã lưu chỉ tiêu thành công!", "success");
        document.getElementById("modal-field-edit").classList.remove("active");
        if (activeReport) selectReport(activeReport);
        loadStats();
      } else {
        showToast("Lỗi khi lưu: " + (data.error || "Unknown"), "error");
      }
    } catch (e) {
      showToast("Lỗi kết nối máy chủ", "error");
    }
  }

  function initSync() {
    const btn = document.getElementById("btn-sync-all");
    btn.addEventListener("click", async () => {
      if (!confirm("Bạn có muốn quét lại toàn bộ dữ liệu từ mã nguồn (.pck, .prc, .xlsx) không?")) return;

      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Đang quét...</span>';

      try {
        const res = await fetch("/api/sync", { method: "POST" });
        const data = await res.json();
        showToast(`Đồng bộ hoàn tất: ${data.total} chỉ tiêu được xử lý!`, "success");
        loadStats();
        loadReports();
      } catch (e) {
        showToast("Lỗi trong quá trình quét", "error");
      } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="fa-solid fa-arrows-rotate"></i> <span>Đồng bộ từ Mã nguồn</span>';
      }
    });
  }
});
