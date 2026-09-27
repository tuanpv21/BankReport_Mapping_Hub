// BankReport Mapping Hub - Master-Detail Modern Frontend Logic

document.addEventListener("DOMContentLoaded", () => {
  let allReports = [];
  let currentSystemFilter = "ALL";
  let activeReport = null;
  let activeReportFields = [];
  let activeTab = "tab-dashboard";

  const tabLinks = document.querySelectorAll(".nav-link");
  const tabViews = document.querySelectorAll(".tab-view");

  initNav();
  initGlobalSearch();
  initExplorer();
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
      chip.className = "report-chip-item";
      chip.innerHTML = `<span>${r.report_code}</span> <span class="chip-count">(${r.field_count})</span>`;
      chip.title = `${r.report_name} - ${r.field_count} chỉ tiêu`;
      chip.addEventListener("click", () => {
        selectReport(r);
        switchTab("tab-explorer");
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
            <p title="${r.report_name}">${r.report_name || r.report_code}</p>
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
    document.getElementById("active-report-desc").textContent = `Bảng đích: ${report.target_table || '-'} | Quy định: ${report.regulation_ref || 'QĐ NHNN'}`;

    document.getElementById("field-filter-input").value = "";
    document.getElementById("filter-field-status").value = "ALL";

    const tbody = document.getElementById("fields-table-body");
    tbody.innerHTML = '<tr><td colspan="7" class="text-center py-5 text-muted"><i class="fa-solid fa-spinner fa-spin"></i> Đang tải dữ liệu chỉ tiêu...</td></tr>';

    try {
      const res = await fetch(`/api/mappings?system=${report.system_code}&report=${report.report_code}`);
      activeReportFields = await res.json();
      filterAndRenderFields();
    } catch (e) {
      tbody.innerHTML = '<tr><td colspan="7" class="text-center py-4 text-rose">Lỗi khi tải chỉ tiêu của báo cáo này.</td></tr>';
    }
  }

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
        (f.transformation_rule && f.transformation_rule.toLowerCase().includes(searchVal));
      return matchStatus && matchSearch;
    });

    document.getElementById("active-fields-count").textContent = `Hiển thị: ${filtered.length} / ${activeReportFields.length} chỉ tiêu`;

    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" class="text-center py-5 text-muted">Không tìm thấy chỉ tiêu nào phù hợp với từ khóa tìm kiếm.</td></tr>';
      return;
    }

    tbody.innerHTML = "";
    filtered.forEach(f => {
      const tr = document.createElement("tr");

      let sourceHtml = "";
      if (f.source_table) {
        sourceHtml = `
          <div class="source-badge-box">
            <span class="src-table-tag"><i class="fa-solid fa-database"></i> ${f.source_table}</span>
            <span class="src-col-tag">${f.source_column || '-'}</span>
          </div>
        `;
      } else {
        sourceHtml = '<span class="unmapped-tag"><i class="fa-solid fa-triangle-exclamation"></i> Chưa map</span>';
      }

      let ruleHtml = "-";
      if (f.transformation_rule) {
        ruleHtml = `<span class="rule-code-snippet" title="${f.transformation_rule}">${truncate(f.transformation_rule, 40)}</span>`;
      } else if (f.notes) {
        ruleHtml = `<span class="text-subtle" style="font-size:12px;">${truncate(f.notes, 35)}</span>`;
      }

      const statusClass = f.status === "Active" ? "status-active" : "status-draft";

      tr.innerHTML = `
        <td><span class="field-code-badge">${f.field_code}</span></td>
        <td>
          <div class="field-name-title">${f.field_name_vi || f.field_code}</div>
          <div class="field-name-sub">${f.lookup_ref ? `<span class="tag-lookup"><i class="fa-solid fa-diagram-next"></i> ${f.lookup_ref}</span>` : ''}</div>
        </td>
        <td><span class="font-mono text-subtle" style="font-size:12px;">${f.data_type || 'VARCHAR2(50)'}</span></td>
        <td>${sourceHtml}</td>
        <td>${ruleHtml}</td>
        <td><span class="status-badge ${statusClass}">${f.status || 'Active'}</span></td>
        <td>
          <div class="action-btn-group">
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

  function truncate(str, max = 40) {
    if (!str) return "";
    return str.length > max ? str.substring(0, max) + "..." : str;
  }

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
            <td><strong class="font-mono text-main">${r.source_table}</strong></td>
            <td><span class="font-mono text-subtle">${r.source_column || '-'}</span></td>
            <td>${sysBadge}</td>
            <td><strong>${r.report_code}</strong></td>
            <td><span class="field-code-badge">${r.field_code}</span></td>
            <td><strong style="color:var(--text-main);">${r.field_name_vi || '-'}</strong></td>
            <td><span class="font-mono text-subtle" style="font-size:11.5px;">${r.implemented_by || '-'}</span></td>
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
          <span class="field-code-badge">${f.field_code}</span> 
          <strong style="font-size:14px; margin-left:8px;">${f.field_name_vi || '-'}</strong>
        </div>

        <div class="detail-label-txt">Báo cáo & Bảng đích:</div>
        <div class="detail-val-txt">
          Mẫu <strong>${f.report_code}</strong> | Bảng: <span class="font-mono text-primary">${f.target_table || '-'}</span>
        </div>

        <div class="detail-label-txt">Kiểu dữ liệu:</div>
        <div class="detail-val-txt font-mono">${f.data_type || 'VARCHAR2(50)'}</div>

        <div class="detail-label-txt">Nguồn Dữ liệu:</div>
        <div class="detail-val-txt">
          ${f.source_table ? `<strong class="font-mono text-primary">${f.source_table}</strong> . <span class="font-mono">${f.source_column || '-'}</span>` : '<span class="unmapped-tag"><i class="fa-solid fa-triangle-exclamation"></i> Chưa map nguồn</span>'}
        </div>

        <div class="detail-label-txt">Quy tắc / Công thức:</div>
        <div class="detail-val-txt">
          <div class="code-box-light">${f.transformation_rule || '(Gán trực tiếp hoặc chưa ghi nhận quy tắc)'}</div>
        </div>

        <div class="detail-label-txt">Bảng mã quy đổi:</div>
        <div class="detail-val-txt">${f.lookup_ref || 'Không sử dụng bảng mã'}</div>

        <div class="detail-label-txt">Thủ tục thực thi:</div>
        <div class="detail-val-txt font-mono">${f.implemented_by || 'Chưa ghi nhận'}</div>

        <div class="detail-label-txt">Căn cứ quy định:</div>
        <div class="detail-val-txt">${f.regulatory_ref || '-'}</div>

        <div class="detail-label-txt">Ghi chú điều kiện:</div>
        <div class="detail-val-txt">${f.notes || '-'}</div>
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