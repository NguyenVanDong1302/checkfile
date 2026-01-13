let lastResult = null;

const LS_KEY = "word_checker_phrases_v1";

function $(id) {
  return document.getElementById(id);
}

function setStatus(msg, type = "muted") {
  const el = $("status");
  el.className = "status " + type;
  el.textContent = msg || "";
}

function escapeHtml(s) {
  return (s ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;",
      }[c])
  );
}

function whereToText(where) {
  if (!where) return "N/A";
  const area = where.area || "body";
  const kind = where.kind || "paragraph";

  if (kind === "paragraph") {
    const p = where.paragraph_index ?? "";
    const page = where.page_est == null ? "?" : where.page_est;
    const line = where.line_est == null ? "?" : where.line_est;
    return `${area} • paragraph #${p} • trang~ ${page} • dòng~ ${line}`;
  }

  if (kind === "table_cell") {
    const t = where.table_index ?? "?";
    const r = where.row ?? "?";
    const c = where.col ?? "?";
    const page = where.page_est == null ? "?" : where.page_est;
    const line = where.line_est == null ? "?" : where.line_est;
    return `${area} • table #${t} (r${r} c${c}) • trang~ ${page} • dòng~ ${line}`;
  }

  const sec = where.section ?? "?";
  const p = where.paragraph_index ?? "?";
  const line = where.line_est == null ? "?" : where.line_est;
  return `${area} • section #${sec} • paragraph #${p} • dòng~ ${line}`;
}

function savePhrasesToLS(text) {
  try {
    localStorage.setItem(LS_KEY, text || "");
  } catch {}
}

function loadPhrasesFromLS() {
  try {
    return localStorage.getItem(LS_KEY) || "";
  } catch {
    return "";
  }
}

function render(result, filterText = "") {
  lastResult = result;

  const stats = result?.stats || {};
  $("summary").innerHTML = `
    <span class="pill">Tổng cụm: <b>${stats.phrases_total ?? 0}</b></span>
    <span class="pill">Tìm thấy: <b class="ok">${stats.found ?? 0}</b></span>
    <span class="pill">Thiếu: <b class="bad">${stats.missing ?? 0}</b></span>
    <span class="pill">Nghi lỗi chính tả: <b>${
      stats.typo_suspects_phrases ?? 0
    }</b></span>
  `;

    const alerts = result?.alerts || {};
    const miss = (alerts.misspellings || []).slice(0, 8);
    const viol = (alerts.rule_violations || []).slice(0, 8);

    let alertsHtml = "";

    if ((alerts.misspellings_total || 0) > 0) {
      alertsHtml += `
      <div class="alertBox">
        <div class="alertTitle bad">⚠️ Từ sai chính tả phát hiện: ${
          alerts.misspellings_total
        }</div>
        ${miss
          .map(
            (x) => `
          <div class="alertItem">
            <div class="muted">${escapeHtml(whereToText(x.where))}</div>
            <div>Sai: <code>${escapeHtml(
              x.wrong_word
            )}</code> (chuẩn/khuyến nghị: <code>${escapeHtml(
              x.canonical
            )}</code>)</div>
            <div>Ngữ cảnh: <code>${escapeHtml(x.snippet || "")}</code></div>
          </div>
        `
          )
          .join("")}
        ${
          (alerts.misspellings_total || 0) > 8
            ? `<div class="muted">... và ${
                alerts.misspellings_total - 8
              } cảnh báo khác.</div>`
            : ""
        }
      </div>
    `;
    }

    if ((alerts.rule_violations_total || 0) > 0) {
      alertsHtml += `
      <div class="alertBox">
        <div class="alertTitle bad">⚠️ Vi phạm quy tắc định dạng: ${
          alerts.rule_violations_total
        }</div>
        ${viol
          .map(
            (x) => `
          <div class="alertItem">

            <div>Ngữ cảnh: <code>${escapeHtml(x.snippet || "")}</code></div>
          </div>
        `
          )
          .join("")}
        ${
          (alerts.rule_violations_total || 0) > 8
            ? `<div class="muted">... và ${
                alerts.rule_violations_total - 8
              } cảnh báo khác.</div>`
            : ""
        }
      </div>
    `;
    }

    $("alerts").innerHTML =
      alertsHtml ||
      `<div class="muted">Không có cảnh báo rule/blacklist.</div>`;


  const hits = result?.hits || [];
  const q = (filterText || "").trim().toLowerCase();

  const filtered = q
    ? hits.filter((h) => (h.phrase || "").toLowerCase().includes(q))
    : hits;

  const tbody = $("tbody");
  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="4" class="muted">Không có kết quả theo bộ lọc.</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered
    .map((h) => {
      const found = !!h.found;
      const count = h.count ?? 0;

      let detailHtml = "";

      const typo = (h.typo_suspects || []).slice(0, 8);

      if (!found) {
        detailHtml = `<div class="muted">Không tìm thấy trong tài liệu.</div>`;
        if (typo.length) {
          detailHtml += `
          <div class="detail" style="margin-top:8px;">
            <div class="muted">⚠️ Nghi lỗi gõ sai gần giống:</div>
            ${typo
              .map(
                (t) => `
              <div class="kv">
                <div class="meta">${escapeHtml(whereToText(t.where))}</div>
                <div>Trong file có: <code>${escapeHtml(
                  t.typed_in_doc
                )}</code> • distance=${t.distance}</div>
                <div>Ngữ cảnh: <code>${escapeHtml(t.snippet || "")}</code></div>
              </div>
            `
              )
              .join("")}
            ${
              (h.typo_suspects || []).length > 8
                ? `<div class="muted">... và ${
                    (h.typo_suspects || []).length - 8
                  } gợi ý khác.</div>`
                : ""
            }
          </div>
        `;
        }
      } else {
        const occs = (h.occurrences || []).slice(0, 10);
        detailHtml = `
        <div class="detail">
          ${occs
            .map(
              (o) => `
            <div class="kv">
              <div class="meta">${escapeHtml(whereToText(o.where))}</div>
              <div>Ngữ cảnh: <code>${escapeHtml(o.snippet || "")}</code></div>
            </div>
          `
            )
            .join("")}
          ${
            count > 10
              ? `<div class="muted">... và ${
                  count - 10
                } vị trí khác (đã ẩn bớt).</div>`
              : ""
          }
        </div>
      `;
      }

      return `
      <tr>
        <td><b>${escapeHtml(h.phrase || "")}</b></td>
        <td>${
          found
            ? `<span class="ok">FOUND</span>`
            : `<span class="bad">MISSING</span>`
        }</td>
        <td>${count}</td>
        <td>${detailHtml}</td>
      </tr>
    `;
    })
    .join("");

  $("btnDownloadJson").disabled = false;
}

function downloadJson(data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "word-check-result.json";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function doCheck() {
  const phrases = $("phrases").value || "";
  const file = $("file").files?.[0];

  if (!file) {
    setStatus("Bạn chưa chọn file .docx hoặc .doc", "bad");
    return;
  }
  if (!phrases.trim()) {
    setStatus("Bạn chưa nhập danh sách cụm từ cần check.", "bad");
    return;
  }

  // Save to localStorage immediately for convenience
  savePhrasesToLS(phrases);

  const fd = new FormData();
  fd.append("file", file);
  fd.append("phrases", phrases);
  fd.append("case_sensitive", $("caseSensitive").checked ? "true" : "false");
  fd.append("whole_word", $("wholeWord").checked ? "true" : "false");
  fd.append("scan_headers_footers", $("scanHF").checked ? "true" : "false");

  fd.append("spellcheck_vi", $("spellcheckVi").checked ? "true" : "false");
  fd.append("spell_max_distance", String($("spellMaxDist").value || 2));
fd.append(
  "check_du_toan_rule",
  $("checkDuToanRule").checked ? "true" : "false"
);

  setStatus("Đang kiểm tra...", "muted");
  $("btnCheck").disabled = true;

  try {
    const res = await fetch("/api/check", { method: "POST", body: fd });
    const data = await res.json();

    if (!res.ok) {
      setStatus(
        (data?.error || "Có lỗi") + (data?.detail ? `: ${data.detail}` : ""),
        "bad"
      );
      return;
    }

    setStatus("Xong.", "ok");
    render(data, $("search").value);
  } catch (e) {
    console.error(e);
    setStatus(
      "Không gọi được server. Hãy mở web qua http://127.0.0.1:8000",
      "bad"
    );
  } finally {
    $("btnCheck").disabled = false;
  }
}

window.addEventListener("DOMContentLoaded", () => {
  // Load phrases from localStorage on startup
  const saved = loadPhrasesFromLS();
  if (saved && !$("phrases").value.trim()) $("phrases").value = saved;

  $("btnCheck").addEventListener("click", doCheck);

  $("phrases").addEventListener("input", (e) => {
    // auto-save while typing (optional)
    savePhrasesToLS(e.target.value || "");
  });

  $("search").addEventListener("input", (e) =>
    render(lastResult || { hits: [], stats: {} }, e.target.value)
  );

  $("btnDownloadJson").addEventListener("click", () => {
    if (lastResult) downloadJson(lastResult);
  });

  $("btnClearSaved").addEventListener("click", () => {
    savePhrasesToLS("");
    $("phrases").value = "";
    setStatus("Đã xoá cụm từ đã lưu.", "ok");
  });
});
