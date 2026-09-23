/* 学生劳务费发放助手 —— 前端逻辑
 *
 * 界面结构：顶栏 + 三步步骤条 + 当前步骤面板 + 发放信息抽屉
 * 录入方式：行内直接改 / 点金额按钮批量设 / 从 Excel 粘一列 / 回车上下跳
 */
(function () {
  "use strict";

  var DEFAULT_FEE_PRESETS = ["500", "800", "1000", "1200", "1500", "1600", "2000"];

  /** 月份预设：只作为兜底（后端读不到时用）。
   *  正常情况用后端给的 periodStart / periodPresets，避免两边各写一套。 */
  function buildPeriodPresets(fromText) {
    var m = /^\s*(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*月?\s*$/.exec(String(fromText || ""));
    var startYear, startMonth;
    if (m) {
      startYear = parseInt(m[1], 10);
      startMonth = parseInt(m[2], 10);
    } else {
      var now = new Date();
      startYear = now.getFullYear();
      startMonth = now.getMonth() + 1;
    }
    var out = [];
    for (var i = 0; i < 6; i++) {
      var y = startYear + Math.floor((startMonth - 1 + i) / 12);
      var mo = (startMonth - 1 + i) % 12 + 1;
      out.push(y + "年" + mo + "月");
    }
    return out;
  }

  // 名单里没填学院时用它兜底（和 store.py 的 DEFAULT_COLLEGE 保持一致）
  var DEFAULT_COLLEGE = "国家网络安全学院";

  // 表格列的录入属性
  var FIELDS = ["studentId", "name", "college", "rate", "hours", "amount"];
  var NUMERIC_FIELDS = { rate: 1, hours: 1, amount: 1 };

  var state = {
    step: 1,
    settings: { unitName: "", projectCode: "", period: "", projectName: "", note: "", projectType: "research" },
    batch: { college: "", rate: "100", hours: "10" },
    feePresets: DEFAULT_FEE_PRESETS.slice(),
    periodPresets: buildPeriodPresets(),
    periodStart: "",        // 后端的月份基准，例如 "2026年6月"
    students: [],
    search: "",
    templates: [],
    tplId: "",
    pending: null,
    pasteTarget: null,
    saveTimer: null
  };

  var seq = 1;
  var rowCache = new Map();

  // ============================================================ 小工具
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) {
    return Array.prototype.slice.call((root || document).querySelectorAll(sel));
  }
  function uid() { return "s" + (seq++) + "_" + Date.now().toString(36); }

  function toNum(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return value;
    var text = String(value).replace(/,/g, "").trim();
    if (text === "") return NaN;
    return /^-?\d+(\.\d+)?$/.test(text) ? parseFloat(text) : NaN;
  }

  function fmt(value) {
    var n = toNum(value);
    if (isNaN(n)) return value === null || value === undefined ? "" : String(value);
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    return String(parseFloat(n.toFixed(2)));
  }

  function money(value) {
    var n = toNum(value);
    if (isNaN(n)) return "0";
    return n.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
  }

  function esc(text) {
    return String(text === null || text === undefined ? "" : text)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function calcAmount(student) {
    if (student.manual) return toNum(student.amount) || 0;
    var rate = toNum(student.rate), hours = toNum(student.hours);
    if (isNaN(rate) || isNaN(hours)) return 0;
    return rate * hours;
  }

  function effAmount(student) {
    if (student.manual) return student.amount === "" ? "" : fmt(student.amount);
    var value = calcAmount(student);
    return value ? fmt(value) : "";
  }

  function checkedStudents() {
    return state.students.filter(function (s) { return s.checked; });
  }

  function studentById(id) {
    for (var i = 0; i < state.students.length; i++) {
      if (state.students[i].id === id) return state.students[i];
    }
    return null;
  }

  function newStudent(seed) {
    seed = seed || {};
    var given = String(seed.college === null || seed.college === undefined
      ? "" : seed.college).trim();
    var fallback = given || state.batch.college || DEFAULT_COLLEGE;
    return {
      id: seed.id || uid(),
      studentId: seed.studentId || "",
      name: seed.name || "",
      // 学院三级兜底：名单里的值 -> 抽屉默认值 -> 全局默认学院
      college: fallback,
      // 这个学院是"兜底填的"还是"名单里本来就有"，界面上要能看出来
      collegeAuto: !given,
      rate: seed.rate || state.batch.rate || "",
      hours: seed.hours || (state.batch.hours === "0" ? "" : (state.batch.hours || "")),
      amount: seed.amount || "",
      manual: !!seed.manual,
      checked: seed.checked !== false
    };
  }

  function normalizeStudent(raw) {
    var amount = raw.amount === null || raw.amount === undefined ? "" : String(raw.amount);
    var manual = (raw.manual === undefined || raw.manual === null)
      ? amount.trim() !== ""
      : !!raw.manual;
    var student = newStudent({
      id: raw.id, studentId: raw.studentId, name: raw.name, college: raw.college,
      rate: raw.rate, hours: raw.hours, amount: amount, manual: manual,
      checked: raw.checked
    });
    // 本地存过的标记优先（用户可能已经知道这是默认值）
    if (raw.collegeAuto !== undefined && raw.collegeAuto !== null) {
      student.collegeAuto = !!raw.collegeAuto;
    }
    return student;
  }

  // ============================================================ 兼职时段
  function renderPeriod() {
    var list = state.periodPresets || [];
    var current = String(state.settings.period || "").trim();

    var bar = $("#periodChips");
    if (bar) {
      bar.classList.toggle("is-empty", list.length === 0);
      bar.innerHTML = list.map(function (value) {
        var active = value === current ? " is-active" : "";
        return '<button type="button" class="period-chip' + active +
          '" data-period="' + esc(value) + '" title="点一下设为「' + esc(value) + '」">' +
          esc(value) + "</button>";
      }).join("");
    }

    var custom = $("#periodCustom");
    if (custom && document.activeElement !== custom) custom.value = current;

    var dl = $("#periodList");
    if (dl) {
      dl.innerHTML = list.map(function (v) {
        return '<option value="' + esc(v) + '"></option>';
      }).join("");
    }

    var hint = $("#periodHint");
    if (hint) {
      hint.textContent = current
        ? ("当前：" + current + "（会写进表头和文件名）")
        : "还没选时段，导出的文件名会没有月份";
    }
  }

  function setPeriod(value) {
    var text = String(value || "").trim();
    state.settings.period = text;
    renderPeriod();
    renderSettings();
    scheduleSave();
    if (text) toast("兼职时段已设为 " + text, "ok");
  }

  function parsePeriodPresets(text) {
    var seen = Object.create(null);
    return String(text || "")
      .split(/[,，、;；\n\r]+/)
      .map(function (t) { return t.trim(); })
      .filter(function (t) {
        if (!t || seen[t]) return false;
        seen[t] = true;
        return true;
      });
  }

  // ============================================================ 提示 / 保存
  function toast(message, kind) {
    var wrap = $("#toastWrap");
    var node = document.createElement("div");
    node.className = "toast " + (kind || "");
    node.textContent = message;
    wrap.appendChild(node);
    setTimeout(function () {
      node.style.transition = "opacity .25s";
      node.style.opacity = "0";
      setTimeout(function () { node.remove(); }, 260);
    }, kind === "err" ? 5200 : 2800);
  }

  function setSaveState(text, kind) {
    var el = $("#saveState");
    el.textContent = text;
    el.className = "save-state " + (kind || "");
  }

  function payload() {
    return {
      step: state.step,
      settings: state.settings,
      batch: state.batch,
      feePresets: state.feePresets,
      periodPresets: state.periodPresets,
      students: state.students
    };
  }

  function scheduleSave() {
    if (state.saveTimer) clearTimeout(state.saveTimer);
    setSaveState("编辑中…");
    state.saveTimer = setTimeout(function () {
      postJSON("/api/settings", payload()).then(function () {
        setSaveState("已保存 " + new Date().toLocaleTimeString("zh-CN", { hour12: false }), "ok");
      }).catch(function (err) {
        setSaveState("保存失败：" + err.message, "err");
      });
    }, 450);
  }

  function saveNow() {
    if (state.saveTimer) clearTimeout(state.saveTimer);
    return postJSON("/api/settings", payload()).then(function () {
      setSaveState("已保存", "ok");
    }).catch(function () { /* 忽略 */ });
  }

  // ============================================================ 接口
  function api(path, options) {
    return fetch(path, options).then(function (res) {
      return res.json().catch(function () { return {}; }).then(function (data) {
        if (!res.ok || data.error) throw new Error(data.error || ("请求失败 " + res.status));
        return data;
      });
    });
  }

  function postJSON(path, body) {
    return api(path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body || {})
    });
  }

  // ============================================================ 步骤
  function gotoStep(step, options) {
    options = options || {};
    state.step = step;
    $$("[data-panel]").forEach(function (panel) {
      panel.hidden = Number(panel.dataset.panel) !== step;
    });
    $$("#steps .step").forEach(function (btn) {
      var no = Number(btn.dataset.step);
      btn.classList.toggle("is-active", no === step);
      btn.classList.toggle("is-done", no < step && stepReady(no));
    });
    if (!options.keepScroll) {
      try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch (err) { /* 忽略 */ }
    }
    renderStepInfo();
    fitTableHeight();
    // 记住停在哪一步，刷新/重启后接着来
    if (!options.silent) scheduleSave();
  }

  function stepReady(no) {
    if (no === 1) return state.students.length > 0;
    if (no === 2) return checkedStudents().length > 0;
    return false;
  }

  function renderStepInfo() {
    var total = state.students.length;
    var picked = checkedStudents().length;
    $("#step1Info").textContent = total ? ("已导入 " + total + " 人") : "还没有学生";
    $("#step2Info").textContent = total
      ? ("已勾选 " + picked + " / " + total + " 人")
      : "先导入名单";

    var sum = 0;
    checkedStudents().forEach(function (s) { sum += calcAmount(s); });
    $("#step3Info").textContent = picked
      ? (picked + " 人 · 合计 " + money(sum) + " 元")
      : "Word 明细表 / Excel 名单";
  }

  // ============================================================ 表格
  function visibleStudents() {
    var q = state.search.trim().toLowerCase();
    if (!q) return state.students;
    return state.students.filter(function (s) {
      return (s.name + " " + s.studentId + " " + s.college).toLowerCase().indexOf(q) >= 0;
    });
  }

  function buildRow(student) {
    var tr = document.createElement("tr");
    tr.dataset.id = student.id;
    tr.innerHTML =
      '<td class="col-check"><input type="checkbox" data-f="checked"></td>' +
      '<td class="col-no" data-no></td>' +
      '<td class="col-sid"><input type="text" class="sid" data-f="studentId" placeholder="学号"></td>' +
      '<td class="col-name"><input type="text" data-f="name" placeholder="姓名"></td>' +
      '<td class="col-college"><input type="text" data-f="college" placeholder="学院"></td>' +
      '<td class="col-rate"><input type="number" class="num" data-f="rate" min="0" step="1" placeholder="0"></td>' +
      '<td class="col-hours"><input type="number" class="num" data-f="hours" min="0" step="1" placeholder="0"></td>' +
      '<td class="col-amount"><input type="number" class="num" data-f="amount" min="0" step="100" placeholder="自动"></td>' +
      '<td class="col-act"><button class="row-del" type="button" title="删除这一行">×</button></td>';
    return tr;
  }

  function fillRow(tr, student, index) {
    if (tr.dataset.checked !== String(student.checked)) {
      tr.dataset.checked = String(student.checked);
      tr.classList.toggle("is-unchecked", !student.checked);
      $('[data-f="checked"]', tr).checked = student.checked;
    }
    $("[data-no]", tr).textContent = String(index + 1);

    $$("input[data-f]", tr).forEach(function (input) {
      var field = input.dataset.f;
      if (field === "checked") return;
      if (document.activeElement === input) return;

      if (field === "amount") {
        var value = effAmount(student);
        if (input.value !== value) input.value = value;
        input.classList.toggle("manual", !!student.manual);
        input.placeholder = student.manual ? "" : "自动";
        input.title = student.manual
          ? "手动填写（点「按标准×工时」可恢复自动计算）"
          : "自动计算：标准 × 工时";
      } else if (field === "college") {
        var college = student.college === null || student.college === undefined
          ? "" : String(student.college);
        if (input.value !== college) input.value = college;
        // 兜底填的学院：淡灰斜体 + 提示，和手填/名单里的区分开
        input.classList.toggle("is-default", !!student.collegeAuto);
        input.title = student.collegeAuto
          ? "名单里没填学院，这里是按默认值自动补的：" + college +
            "（改一下就会变成你自己的值）"
          : college;
      } else {
        var raw = student[field] === null || student[field] === undefined
          ? "" : String(student[field]);
        if (input.value !== raw) input.value = raw;
        // 列宽有限，长名字/长学院名会被输入框截断——挂个 title，鼠标停上去能看全
        input.title = raw;
      }
    });
  }

  function renderTable() {
    var tbody = $("#tbody");
    var rows = visibleStudents();
    var seen = new Set();

    rows.forEach(function (student, index) {
      var tr = rowCache.get(student.id);
      if (!tr) {
        tr = buildRow(student);
        rowCache.set(student.id, tr);
        tbody.appendChild(tr);
      }
      seen.add(student.id);
      fillRow(tr, student, index);
    });

    rowCache.forEach(function (tr, id) {
      if (!seen.has(id)) { tr.remove(); rowCache.delete(id); }
    });

    var order = new Map(rows.map(function (s, i) { return [s.id, i]; }));
    $$("tr", tbody)
      .sort(function (a, b) {
        return (order.get(a.dataset.id) || 0) - (order.get(b.dataset.id) || 0);
      })
      .forEach(function (tr) { tbody.appendChild(tr); });

    $("#empty").hidden = state.students.length > 0;
    if (state.students.length === 0 && rowCache.size) {
      tbody.innerHTML = "";
      rowCache.clear();
    }

    renderSummary();
    renderImported();
    renderStepInfo();
    refreshTableHeight();
  }

  function renderSummary() {
    var all = state.students;
    var picked = checkedStudents();
    var totalAmount = 0, totalHours = 0;
    picked.forEach(function (s) {
      totalAmount += calcAmount(s);
      var h = toNum(s.hours);
      if (!isNaN(h)) totalHours += h;
    });

    $("#summary").innerHTML =
      '<span class="item">已勾选 <b>' + picked.length + "</b> 人</span>" +
      '<span class="item">总工时 <b>' + money(totalHours) + "</b></span>" +
      '<span class="item total">实发合计 <b>' + money(totalAmount) + " 元</b></span>" +
      '<span class="grow"></span>' +
      '<span class="item">未勾选 ' + (all.length - picked.length) + " 人</span>";
  }

  function renderImported() {
    var box = $("#importedBox");
    var list = state.students;
    box.hidden = list.length === 0;
    if (!list.length) return;
    $("#importedSummary").textContent = "共 " + list.length + " 人";
    $("#importedChips").innerHTML = list.map(function (s) {
      return '<span class="chip">' + esc(s.name || "（未填姓名）") +
        (s.studentId ? ' <b>' + esc(s.studentId) + "</b>" : "") + "</span>";
    }).join("");
  }

  // ============================================================ 侧栏渲染
  function renderSettings() {
    $$("[data-setting]").forEach(function (el) {
      if (document.activeElement === el) return;
      var key = el.dataset.setting;
      var value = state.settings[key] === null || state.settings[key] === undefined
        ? "" : String(state.settings[key]);
      if (el.value !== value) el.value = value;
    });
    $$("[data-batch]").forEach(function (el) {
      if (document.activeElement === el) return;
      var key = el.dataset.batch;
      var value = state.batch[key] === null || state.batch[key] === undefined
        ? "" : String(state.batch[key]);
      if (el.value !== value) el.value = value;
    });
    $("#batchRate").placeholder = state.batch.rate || "100";
    $("#batchHours").placeholder = state.batch.hours || "10";
    $("#batchCollege").placeholder = state.batch.college || "学院";
  }

  function renderPresets() {
    var list = state.feePresets || [];
    var bar = $("#presetBar");
    bar.classList.toggle("is-empty", list.length === 0);
    bar.innerHTML = list.map(function (value) {
      return '<button type="button" class="preset-btn" data-preset="' + esc(value) +
        '" title="给勾选的学生设为 ' + esc(value) + ' 元">' + esc(value) + "</button>";
    }).join("");

    syncPresetModal();
  }

  /** 预设弹窗两种用途共用：补贴费 / 兼职时段。内容按当前用途同步。 */
  function syncPresetModal() {
    var modal = $("#modalPreset");
    if (!modal || modal.hidden) return;
    var input = $("#presetInputModal");
    if (!input || document.activeElement === input) return;
    input.value = state.modalPresetKind === "period"
      ? (state.periodPresets || []).join(", ")
      : (state.feePresets || []).join(", ");
  }

  function openPresetModal(kind) {
    state.modalPresetKind = kind === "period" ? "period" : "fee";
    var isPeriod = state.modalPresetKind === "period";
    $("#presetModalTitle").textContent = isPeriod ? "兼职时段预设" : "补贴费预设";
    $("#presetModalHint").textContent = isPeriod
      ? "逗号或换行分隔。这些月份会变成第 1 步顶部的快捷按钮，点一下就能选用。"
      : "逗号或空格分隔，可填小数。这些按钮会出现在第 2 步顶部。";
    var input = $("#presetInputModal");
    input.placeholder = isPeriod ? "2026年9月, 2026年10月, 2026年11月"
                                 : "500, 800, 1000, 1200, 1500, 1600, 2000";
    input.value = isPeriod
      ? (state.periodPresets || []).join(", ")
      : (state.feePresets || []).join(", ");
    $("#modalPreset").hidden = false;
    input.focus();
  }

  function renderTemplates() {
    var select = $("#tplSelect");
    var list = state.templates;
    select.innerHTML = '<option value="">— 未选择 —</option>' +
      list.map(function (t) {
        var meta = [t.period, t.studentCount ? t.studentCount + "人" : ""]
          .filter(Boolean).join(" · ");
        return '<option value="' + esc(t.id) + '">' + esc(t.name) +
          (meta ? "（" + esc(meta) + "）" : "") + "</option>";
      }).join("");
    select.value = state.tplId || "";
  }

  function renderRecap() {
    var picked = checkedStudents();
    var sum = 0, hours = 0;
    picked.forEach(function (s) {
      sum += calcAmount(s);
      var h = toNum(s.hours);
      if (!isNaN(h)) hours += h;
    });
    var box = $("#exportRecap");
    box.classList.toggle("is-warn", picked.length === 0);
    var noName = picked.filter(function (s) { return !String(s.name).trim(); }).length;
    box.innerHTML =
      '<span class="item">将导出 <b>' + picked.length + "</b> 人</span>" +
      '<span class="item">总工时 <b>' + money(hours) + "</b></span>" +
      '<span class="item">合计 <b>' + money(sum) + " 元</b></span>" +
      (noName ? '<span class="item">⚠️ <b>' + noName + "</b> 人没填姓名</span>" : "") +
      (picked.length ? "" : '<span class="item">⚠️ 还没有勾选学生，请先回到第 2 步</span>');
  }

  // ============================================================ 发放信息抽屉
  function openDrawer() {
    $("#drawer").hidden = false;
    $("#drawerMask").hidden = false;
  }
  function closeDrawer() {
    $("#drawer").hidden = true;
    $("#drawerMask").hidden = true;
  }

  // ============================================================ 金额操作
  function blurActiveAmount() {
    var el = document.activeElement;
    if (el && el.dataset && el.dataset.f === "amount" && el.blur) el.blur();
  }

  function setAmounts(students, value) {
    students.forEach(function (s) {
      s.amount = String(value);
      s.manual = true;
    });
  }

  function applyPreset(value) {
    var amount = toNum(value);
    if (isNaN(amount)) return;

    // 正在编辑某一行：只改这一行
    var focused = document.activeElement;
    if (focused && focused.dataset && focused.dataset.f === "amount") {
      var tr = focused.closest("tr");
      var student = tr && studentById(tr.dataset.id);
      if (student) {
        student.amount = String(amount);
        student.manual = true;
        focused.value = String(amount);
        focused.classList.add("manual");
        focused.placeholder = "";
        renderSummary(); renderStepInfo(); renderRecap();
        scheduleSave();
        return toast("这一行设为 " + fmt(amount) + " 元", "ok");
      }
    }

    var picked = checkedStudents();
    if (!picked.length) {
      return toast("请先勾选学生；或点进某一行的金额格再点按钮，只改那一行", "warn");
    }
    blurActiveAmount();
    setAmounts(picked, amount);
    renderTable(); renderRecap(); scheduleSave();
    toast("已把 " + picked.length + " 位学生的补贴费设为 " + fmt(amount) + " 元", "ok");
  }

  function autoAmounts() {
    var picked = checkedStudents();
    if (!picked.length) return toast("请先勾选学生", "warn");
    blurActiveAmount();
    picked.forEach(function (s) { s.manual = false; s.amount = ""; });
    renderTable(); renderRecap(); scheduleSave();
    toast("已改回 标准 × 工时 自动计算", "ok");
  }

  // ============================================================ 粘贴一列
  function parseClipboard(text) {
    return String(text || "")
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map(function (line) { return line.replace(/\n$/, ""); })
      .filter(function (line, index, all) {
        return !(index === all.length - 1 && line === "");
      })
      .map(function (line) { return line.split("\t"); });
  }

  /** 从 (rowIndex, field) 开始向下铺一列数据（像 Excel 那样）。 */
  function pasteColumn(startStudentId, field, grid) {
    var flat = [];
    grid.forEach(function (cells) {
      if (cells.length === 1 && cells[0].trim() === "" && grid.length > 1) return;
      cells.forEach(function (cell) { flat.push(String(cell).trim()); });
    });
    if (!flat.length) return { count: 0 };

    var rows = visibleStudents();
    var start = -1;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].id === startStudentId) { start = i; break; }
    }
    if (start < 0) start = 0;

    var applied = 0, skipped = 0;
    for (var k = 0; k < flat.length && start + k < rows.length; k++) {
      var student = rows[start + k];
      if (!student.checked) { skipped++; continue; }
      var value = flat[k];
      if (NUMERIC_FIELDS[field]) {
        var n = toNum(value);
        if (value === "") {
          student[field] = "";
        } else if (isNaN(n)) {
          continue;
        } else {
          student[field] = String(n);
        }
      } else {
        student[field] = value;
      }
      if (field === "amount") student.manual = value !== "";
      applied++;
    }
    return { count: applied, skipped: skipped };
  }

  /** 从 (起始行, 起始列) 开始铺一块矩形数据 —— 和 Excel 一样：
   *  粘 3 列就横着填 3 列（学号/姓名/学院…），每行往下走一行。
   *  未勾选的行会跳过（和单列粘贴保持一致）。
   */
  function pasteBlock(startStudentId, startField, grid) {
    var rows = visibleStudents();
    var startRow = 0;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].id === startStudentId) { startRow = i; break; }
    }

    var colCount = grid[0].length;
    var startCol = FIELDS.indexOf(startField);
    if (startCol < 0) startCol = 0;
    var maxCols = Math.min(colCount, FIELDS.length - startCol);

    var applied = 0, skipped = 0, cells = 0;
    for (var r = 0; r < grid.length && startRow + r < rows.length; r++) {
      var student = rows[startRow + r];
      if (!student.checked) { skipped++; continue; }
      var rowTouched = false;
      for (var c = 0; c < maxCols; c++) {
        var field = FIELDS[startCol + c];
        if (!field || field === "checked") continue;
        var value = String(grid[r][c] === undefined ? "" : grid[r][c]).trim();
        if (NUMERIC_FIELDS[field]) {
          if (value === "") { student[field] = ""; rowTouched = true; cells++; continue; }
          var n = toNum(value);
          if (isNaN(n)) continue;          // 数值列里的非数字跳过，不影响其它列
          student[field] = String(n);
        } else {
          student[field] = value;
        }
        if (field === "amount") student.manual = value !== "";
        rowTouched = true;
        cells++;
      }
      if (rowTouched) applied++;
    }
    return { rows: applied, cells: cells, skipped: skipped, cols: maxCols };
  }

  function handlePaste(event, input) {
    var text = (event.clipboardData || window.clipboardData || {}).getData
      ? (event.clipboardData || window.clipboardData).getData("text")
      : "";
    if (!text) return;

    var grid = parseClipboard(text);
    if (!grid.length || !grid[0]) return;
    var isMulti = grid.length > 1 || grid[0].length > 1;
    if (!isMulti) return;   // 单值：交给浏览器默认处理

    event.preventDefault();
    var tr = input.closest("tr");
    var field = input.dataset.f;

    // 多列 -> 按块铺（Excel 语义）；单列 -> 只往下铺
    var isBlock = grid.some(function (line) { return line.length > 1; });
    if (isBlock) {
      var block = pasteBlock(tr.dataset.id, field, grid);
      if (!block.cells) {
        return toast("没有可填入的内容（未勾选的行会被跳过）", "warn");
      }
      renderTable(); renderRecap(); scheduleSave();
      toast("已按块填入 " + block.rows + " 行 × " + block.cols + " 列" +
        (block.skipped ? "（跳过未勾选 " + block.skipped + " 行）" : ""), "ok");
      return;
    }

    var result = pasteColumn(tr.dataset.id, field, grid);
    if (!result.count) {
      return toast("没有可填入的行（未勾选的行会被跳过）", "warn");
    }
    renderTable(); renderRecap(); scheduleSave();
    toast("已从光标处向下填入 " + result.count + " 行" +
      (result.skipped ? "（跳过未勾选 " + result.skipped + " 行）" : ""), "ok");
  }

  /** 焦点所在列高亮，提示"可以在这里粘一列"。 */
  function markPasteColumn(input) {
    $$("#tbody tr").forEach(function (tr) { tr.classList.remove("is-pastable"); });
    var previous = state.pasteTarget;
    if (previous) {
      $$('#tbody input[data-f="' + previous + '"]').forEach(function (el) {
        el.classList.remove("is-pastable");
      });
    }
    if (!input) { state.pasteTarget = null; return; }
    var field = input.dataset.f;
    if (!field || field === "checked") { state.pasteTarget = null; return; }

    $$('#tbody input[data-f="' + field + '"]').forEach(function (el) {
      el.classList.add("is-pastable");
    });
    state.pasteTarget = field;
  }

  function clearPasteColumn() {
    if (state.pasteTarget) {
      $$('#tbody input[data-f="' + state.pasteTarget + '"]').forEach(function (el) {
        el.classList.remove("is-pastable");
      });
    }
    state.pasteTarget = null;
  }

  /** 回车 / 上下键：同列跳到上/下一行。 */
  function moveFocus(input, delta) {
    var tr = input.closest("tr");
    var field = input.dataset.f;
    var rows = $$("#tbody tr");
    var index = rows.indexOf(tr);
    var next = rows[index + delta];
    if (!next) return false;
    var target = next.querySelector('input[data-f="' + field + '"]');
    if (!target) return false;
    target.focus();
    target.select && target.select();
    return true;
  }

  // ============================================================ 导入
  function mergeStudents(incoming, merge) {
    var added = 0, updated = 0;
    incoming.forEach(function (raw) {
      var fresh = normalizeStudent(raw);
      var hit = merge ? state.students.filter(function (x) {
        return (fresh.studentId && x.studentId === fresh.studentId) ||
               (fresh.name && x.name === fresh.name);
      })[0] : null;

      if (hit) {
        hit.studentId = fresh.studentId || hit.studentId;
        hit.name = fresh.name || hit.name;
        hit.college = fresh.college || hit.college || DEFAULT_COLLEGE;
        if (fresh.rate) hit.rate = fresh.rate;
        if (fresh.hours) hit.hours = fresh.hours;
        if (fresh.amount) { hit.amount = fresh.amount; hit.manual = true; }
        hit.checked = true;
        updated++;
      } else {
        state.students.push(fresh);
        added++;
      }
    });
    return { added: added, updated: updated };
  }

  function showPreview(payloadData) {
    state.pending = payloadData;
    var list = payloadData.students;
    if (!list.length) {
      toast("没有解析到学生，请检查格式", "warn");
      $("#preview").hidden = true;
      return;
    }
    $("#previewTitle").textContent = "解析到 " + list.length + " 位学生 · " + payloadData.source;
    $("#previewList").innerHTML =
      '<table><thead><tr><th>学号</th><th>姓名</th><th>津贴/金额</th><th>状态</th></tr></thead><tbody>' +
      list.slice(0, 200).map(function (s) {
        var hit = state.students.filter(function (x) {
          return (s.studentId && x.studentId === s.studentId) || (s.name && x.name === s.name);
        })[0];
        return '<tr class="' + (hit ? "is-update" : "") + '">' +
          '<td class="mono">' + esc(s.studentId) + "</td>" +
          "<td>" + esc(s.name) + "</td>" +
          "<td>" + esc(s.amount) + "</td>" +
          "<td>" + (hit ? "更新已有" : "新增") + "</td></tr>";
      }).join("") +
      (list.length > 200 ? '<tr><td colspan="4">… 其余 ' + (list.length - 200) + " 位省略</td></tr>" : "") +
      "</tbody></table>";
    $("#preview").hidden = false;
  }

  function confirmImport() {
    if (!state.pending) return;
    var result = mergeStudents(state.pending.students, $("#mergeToggle").checked);
    state.pending = null;
    $("#preview").hidden = true;
    $("#pasteArea").value = "";
    $("#pasteBox").hidden = true;
    renderTable(); saveNow();
    toast("导入完成：新增 " + result.added + " 人，更新 " + result.updated + " 人", "ok");
    if (state.students.length) gotoStep(2);
  }

  function uploadRoster(file) {
    var form = new FormData();
    form.append("file", file);
    api("/api/roster/upload", { method: "POST", body: form }).then(function (data) {
      showPreview({ students: data.students || [], source: data.source || file.name });
    }).catch(function (err) {
      toast("读取文件失败：" + err.message, "err");
    }).then(function () {
      $("#fileRoster").value = "";
    });
  }

  // ============================================================ 模板
  function selectedTemplate() {
    state.tplId = $("#tplSelect").value;
    return state.tplId;
  }

  function applyTemplate() {
    var id = selectedTemplate();
    if (!id) return toast("请先选择一个模板", "warn");
    postJSON("/api/templates/load", { id: id }).then(function (data) {
      var tpl = data.template || {};
      state.settings = Object.assign({}, state.settings, tpl.settings || {});
      state.batch = Object.assign({}, state.batch, tpl.batch || {});
      if (tpl.feePresets && tpl.feePresets.length) state.feePresets = tpl.feePresets;
      if (tpl.periodPresets && tpl.periodPresets.length) state.periodPresets = tpl.periodPresets;
      if (tpl.students && tpl.students.length) {
        state.students = tpl.students.map(normalizeStudent);
        rowCache.clear();
        $("#tbody").innerHTML = "";
      } else {
        state.students.forEach(function (s) {
          if (state.batch.college) s.college = state.batch.college;
          if (state.batch.rate) s.rate = state.batch.rate;
          if (state.batch.hours && state.batch.hours !== "0") s.hours = state.batch.hours;
        });
      }
      renderAll(); saveNow();
      toast("已套用模板「" + (tpl.name || id) + "」", "ok");
    }).catch(function (err) { toast("套用失败：" + err.message, "err"); });
  }

  function saveTemplate(forceId) {
    var id = forceId || "";
    var name;
    if (id) {
      var found = state.templates.filter(function (t) { return t.id === id; })[0];
      name = found ? found.name : id;
    } else {
      name = window.prompt("给这个模板起个名字：", state.settings.period
        ? state.settings.period + " 发放模板" : "发放模板");
      if (name === null) return;
      name = name.trim();
      if (!name) return toast("名字不能为空", "warn");
    }
    var includeStudents = window.confirm(
      "是否把当前 " + state.students.length + " 位学生名单一起存进模板？\n\n" +
      "确定 = 连名单一起记住\n取消 = 只记住发放信息和默认值");

    postJSON("/api/templates/save", {
      id: id, name: name,
      settings: state.settings, batch: state.batch,
      feePresets: state.feePresets, periodPresets: state.periodPresets,
      includeStudents: includeStudents,
      students: includeStudents ? state.students : []
    }).then(function (data) {
      state.templates = data.templates || [];
      state.tplId = data.template.id;
      renderTemplates();
      toast("模板「" + name + "」已保存", "ok");
    }).catch(function (err) { toast("保存失败：" + err.message, "err"); });
  }

  function deleteTemplate() {
    var id = selectedTemplate();
    if (!id) return toast("请先选择一个模板", "warn");
    if (!window.confirm("确定删除模板「" + id + "」吗？")) return;
    postJSON("/api/templates/delete", { id: id }).then(function (data) {
      state.templates = data.templates || [];
      state.tplId = "";
      renderTemplates();
      toast("模板已删除", "ok");
    }).catch(function (err) { toast("删除失败：" + err.message, "err"); });
  }

  function uploadDocx(file) {
    var form = new FormData();
    form.append("file", file);
    form.append("projectType", state.settings.projectType || "research");
    $("#btnUploadDocx").disabled = true;
    api("/api/template/docx", { method: "POST", body: form }).then(function (data) {
      var s = data.settings || {};
      Object.keys(s).forEach(function (k) { if (s[k]) state.settings[k] = s[k]; });
      var added = 0;
      (data.students || []).forEach(function (raw) {
        var hit = state.students.filter(function (x) {
          return (raw.studentId && x.studentId === raw.studentId) ||
                 (raw.name && x.name === raw.name);
        })[0];
        if (hit) {
          ["college", "rate", "hours"].forEach(function (k) { if (raw[k]) hit[k] = raw[k]; });
          if (raw.amount && !hit.manual) { hit.amount = raw.amount; hit.manual = true; }
        } else {
          state.students.push(normalizeStudent(raw));
          added++;
        }
      });
      renderAll(); saveNow();
      toast("已换用 " + data.file + " 作底板" + (added ? "，补充 " + added + " 位学生" : ""), "ok");
    }).catch(function (err) {
      toast("导入模板失败：" + err.message, "err");
    }).then(function () {
      $("#btnUploadDocx").disabled = false;
      $("#fileDocx").value = "";
    });
  }

  // ============================================================ 导出
  function exportFiles(kinds) {
    var picked = checkedStudents();
    if (!picked.length) {
      toast("请先在第 2 步勾选学生", "warn");
      return gotoStep(2);
    }
    if (kinds.indexOf("docx") >= 0) {
      var noName = picked.filter(function (s) { return !String(s.name).trim(); });
      if (noName.length) {
        return toast("有 " + noName.length + " 位学生没填姓名，无法生成明细表", "err");
      }
    }

    var cards = ["#btnExportDocx", "#btnExportXls", "#btnExportBoth"];
    cards.forEach(function (sel) { $(sel).disabled = true; });
    $$(".result").forEach(function (n) { n.remove(); });

    postJSON("/api/export", {
      settings: state.settings, batch: state.batch,
      feePresets: state.feePresets, periodPresets: state.periodPresets,
      students: state.students, kinds: kinds, label: $("#exportLabel").value.trim()
    }).then(function (data) {
      var results = data.results || {};
      var html = [];
      if (results.docx) {
        html.push(resultRow("Word 明细表", results.docx.file, results.docx.path,
          results.docx.students + " 人 · 合计 " + money(results.docx.total) + " 元"));
      }
      if (results.xls) {
        html.push(resultRow("Excel 名单", results.xls.file, results.xls.path,
          results.xls.students + " 人"));
      }
      $("#results").innerHTML = html.join("");
      setSaveState("已导出 " + new Date().toLocaleTimeString("zh-CN", { hour12: false }), "ok");
      toast("导出成功，文件已下载并保存到 output 文件夹", "ok");
    }).catch(function (err) {
      $("#results").innerHTML = resultRow("导出失败", err.message, "", "");
      toast("导出失败：" + err.message, "err");
    }).then(function () {
      cards.forEach(function (sel) { $(sel).disabled = false; });
    });
  }

  function resultRow(title, file, path, meta) {
    var href = "/api/export/download?file=" + encodeURIComponent(file) + "&t=" + Date.now();
    return '<div class="result">' +
      '<span class="tile-icon">✅</span>' +
      '<span class="meta"><strong>' + esc(title) + " · " + esc(file) + "</strong>" +
      "<code>" + esc(path) + "</code></span>" +
      '<span class="tip">' + esc(meta) + "</span>" +
      '<a class="btn btn-ghost" href="' + href + '">重新下载</a>' +
      "</div>";
  }

  /** 把元素滚进视野；不支持该 API 的环境（如 jsdom）静默跳过。 */
  function scrollIntoViewSafe(el, options) {
    if (el && typeof el.scrollIntoView === "function") {
      try { el.scrollIntoView(options || { block: "nearest" }); } catch (err) { /* 忽略 */ }
    }
  }

  // ============================================================ 批量填入一列
  var FILL_LABEL = {
    hours: "工时",
    rate: "标准（元/时）",
    amount: "实发金额"
  };

  function openFillPanel() {
    var panel = $("#fillPanel");
    if (!checkedStudents().length) {
      return toast("请先勾选学生，批量填入只作用于勾选的人", "warn");
    }
    // 按钮文案是"批量填入一列…"，语义就是打开——不做 toggle，
    // 否则面板已开时再点一下会关掉，很容易让人以为按钮坏了。收起用面板里的「取消」。
    panel.hidden = false;
    $("#btnBatchOpen").textContent = "收起批量填入";
    renderFillTotalBox();
    runFillPreview();
    fitTableHeight();
    scrollIntoViewSafe($("#fillInput"));
  }

  function closeFillPanel() {
    $("#fillPanel").hidden = true;
    $("#btnBatchOpen").textContent = "批量填入一列…";
    fitTableHeight();
  }

  function currentFillField() {
    var tab = $(".filltab.is-active");
    return tab ? tab.dataset.fillField : "hours";
  }

  function renderFillTotalBox() {
    // 只有「实发金额」支持填总额均摊
    var isAmount = currentFillField() === "amount";
    $("#fillTotalBox").hidden = !isAmount;
    if (!isAmount) {
      $("#fillUseTotal").checked = false;
      $("#fillTotal").hidden = true;
    } else {
      $("#fillTotal").hidden = !$("#fillUseTotal").checked;
    }
  }

  /** 请求服务端做配对预览（金额由后端按同一口径算，和导出文件一致） */
  function runFillPreview() {
    var field = currentFillField();
    var picked = checkedStudents();
    var preview = $("#fillPreview");
    var applyBtn = $("#btnFillApply");

    if (!picked.length) {
      preview.innerHTML = '<div class="fillpanel-warn">还没有勾选学生，请先在上面打勾</div>';
      applyBtn.disabled = true;
      return;
    }

    var useTotal = field === "amount" && $("#fillUseTotal").checked;
    var body = {
      field: field,
      students: state.students,
      values: $("#fillInput").value.split("\n")
    };
    if (useTotal) body.total = $("#fillTotal").value;

    if (!useTotal && !$("#fillInput").value.trim()) {
      // 不要静默把按钮置灰，直接告诉用户下一步做什么
      preview.innerHTML = '<div class="fillpanel-hint">' +
        (field === "amount"
          ? "在这里粘贴一列金额，或勾选上面的「填总额，平均分给勾选的人」"
          : "在这里粘贴一列" + esc(FILL_LABEL[field] || field) +
            "，粘贴后这里会显示谁拿到多少、一共多少") +
        "</div>";
      applyBtn.disabled = true;
      state.fillPlan = null;
      return;
    }

    postJSON("/api/batch/preview", body).then(function (data) {
      renderFillPreview(data, picked);
      applyBtn.disabled = !data.count;
      state.fillPlan = data;
    }).catch(function (err) {
      preview.innerHTML = '<div class="fillpanel-warn">' + esc(err.message) + "</div>";
      applyBtn.disabled = true;
      state.fillPlan = null;
    });
  }

  function renderFillPreview(data, picked) {
    var field = data.field;
    var valueHead = FILL_LABEL[field] || field;
    var items = data.items || [];

    var rows = items.map(function (item, index) {
      var given = item.value !== null && item.value !== undefined;
      var cls = given ? (item.manual && field !== "amount" ? "is-manual" : "") : "is-skip";
      return '<tr class="' + cls + '">' +
        '<td class="num">' + (index + 1) + "</td>" +
        "<td>" + esc(item.name || "（未填姓名）") + "</td>" +
        '<td class="num">' + (given ? esc(fmt(item.value)) : "— 保持原样 —") + "</td>" +
        '<td class="num amount">' + esc(money(item.amount)) + "</td>" +
        "</tr>";
    }).join("");

    var sum = money(data.totalAmount || 0);
    var head = '<div class="fillpanel-sum">' +
      "<span>已勾选 <b>" + picked.length + "</b> 人</span>" +
      "<span>本次填入 <b>" + (data.count || 0) + "</b> 个" + esc(valueHead) + "</span>" +
      (data.mode === "total"
        ? "<span>每人 <b>" + esc(money(data.unit)) + "</b> 元（总额 " + esc(money(data.total)) + "）</span>"
        : "<span>填入部分小计 <b>" + esc(money(data.filledAmount || 0)) + "</b> 元</span>") +
      (data.keptCount
        ? "<span>另有 " + data.keptCount + " 人未填、保持原样</span>" : "") +
      "<span>填完后全表合计 <b>" + sum + "</b> 元</span>" +
      "</div>";

    $("#fillPreview").innerHTML = head +
      '<table><thead><tr><th class="num">#</th><th>学生</th><th class="num">' +
      esc(valueHead) + '</th><th class="num">实发金额</th></tr></thead><tbody>' +
      rows + "</tbody></table>";
  }

  function applyFill() {
    var plan = state.fillPlan;
    if (!plan || !plan.count) return toast("先填写要填入的数值", "warn");
    var field = plan.field;
    var applied = 0;

    (plan.items || []).forEach(function (item) {
      if (item.value === null || item.value === undefined) return;
      var student = studentById(item.id);
      if (!student) return;
      student[field] = String(item.value);
      if (field === "amount") student.manual = true;
      applied++;
    });

    blurActiveAmount();
    closeFillPanel();
    $("#fillInput").value = "";
    $("#fillUseTotal").checked = false;
    state.fillPlan = null;
    renderTable(); renderRecap(); scheduleSave();
    toast("已给 " + applied + " 位学生填入" + (FILL_LABEL[field] || field), "ok");
  }

  // ============================================================ 批量
  function applyBatch(what) {
    var picked = checkedStudents();
    if (!picked.length) return toast("请先勾选学生（批量只作用于勾选的行）", "warn");

    var input = what === "rate" ? $("#batchRate")
      : what === "hours" ? $("#batchHours") : $("#batchCollege");
    var value = input.value.trim();
    if (!value) return toast("请先填入要套用的值", "warn");

    blurActiveAmount();
    picked.forEach(function (s) { s[what] = value; });
    input.value = "";
    renderTable(); renderRecap(); scheduleSave();
    toast("已给 " + picked.length + " 位学生套用" +
      (what === "rate" ? "标准" : what === "hours" ? "工时" : "学院"), "ok");
  }

  function selectAll(mode) {
    visibleStudents().forEach(function (s) {
      if (mode === "all") s.checked = true;
      else if (mode === "none") s.checked = false;
      else s.checked = !s.checked;
    });
    renderTable(); renderRecap(); scheduleSave();
  }

  // ============================================================ 预设
  function parsePresets(text) {
    var seen = Object.create(null);
    return String(text || "")
      .split(/[,，、;；\s]+/)
      .map(function (t) { return t.trim(); })
      .filter(function (t) {
        if (!t) return false;
        var n = toNum(t);
        if (isNaN(n) || n < 0) return false;
        if (seen[t]) return false;
        seen[t] = true;
        return true;
      })
      .map(function (t) { return fmt(t); });
  }

  // ============================================================ 事件
  function bind() {
    // --- 顶栏 ---
    $("#btnDrawer").addEventListener("click", openDrawer);
    $("#btnDrawerClose").addEventListener("click", closeDrawer);
    $("#drawerMask").addEventListener("click", closeDrawer);
    $("#btnHelp").addEventListener("click", function () { $("#modalHelp").hidden = false; });

    // --- 步骤条 ---
    $$("#steps .step").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var target = Number(btn.dataset.step);
        if (target > 1 && !state.students.length) {
          toast("请先导入学生名单", "warn");
          return gotoStep(1);
        }
        gotoStep(target);
      });
    });
    $$("[data-goto]").forEach(function (btn) {
      btn.addEventListener("click", function () { gotoStep(Number(btn.dataset.goto)); });
    });

    // --- 第1步：三种导入方式 ---
    $("#tileFile").addEventListener("click", function () { $("#fileRoster").click(); });
    $("#drop").addEventListener("click", function () { $("#fileRoster").click(); });
    $("#fileRoster").addEventListener("change", function (e) {
      if (e.target.files[0]) uploadRoster(e.target.files[0]);
    });
    $("#tilePaste").addEventListener("click", function () {
      var box = $("#pasteBox");
      box.hidden = !box.hidden;
      if (!box.hidden) { $("#pasteArea").focus(); scrollIntoViewSafe($("#pasteArea")); }
    });
    $("#tileManual").addEventListener("click", function () {
      addStudent();
      gotoStep(2);
    });

    ["dragenter", "dragover"].forEach(function (evt) {
      document.addEventListener(evt, function (e) {
        if (e.target.closest && e.target.closest(".drop")) {
          e.preventDefault(); $("#drop").classList.add("is-over");
        }
      });
    });
    ["dragleave", "drop"].forEach(function (evt) {
      document.addEventListener(evt, function (e) {
        if (e.target.closest && e.target.closest(".drop")) {
          e.preventDefault(); $("#drop").classList.remove("is-over");
        }
      });
    });
    document.addEventListener("drop", function (e) {
      if (e.target.closest && e.target.closest(".drop") && e.dataTransfer.files[0]) {
        uploadRoster(e.dataTransfer.files[0]);
      }
    });

    $("#btnParseText").addEventListener("click", function () {
      var text = $("#pasteArea").value;
      if (!text.trim()) return toast("请先粘贴名单", "warn");
      postJSON("/api/roster/import", { text: text }).then(function (data) {
        showPreview({ students: data.students || [], source: "粘贴文本" });
      }).catch(function (err) { toast("解析失败：" + err.message, "err"); });
    });
    $("#btnConfirmImport").addEventListener("click", confirmImport);
    $("#btnCancelImport").addEventListener("click", function () {
      state.pending = null;
      $("#preview").hidden = true;
    });
    $("#btnClearStudents").addEventListener("click", function () {
      if (!window.confirm("清空当前 " + state.students.length + " 位学生？")) return;
      state.students = [];
      rowCache.clear();
      $("#tbody").innerHTML = "";
      renderTable(); saveNow();
    });

    // --- 第2步：金额按钮 ---
    $("#presetBar").addEventListener("click", function (e) {
      var btn = e.target.closest("[data-preset]");
      if (btn) applyPreset(btn.dataset.preset);
    });
    $("#btnAmountAuto").addEventListener("click", autoAmounts);

    // --- 批量填入一列 ---
    $("#btnBatchOpen").addEventListener("click", openFillPanel);
    $("#btnFillCancel").addEventListener("click", closeFillPanel);
    $("#btnFillApply").addEventListener("click", applyFill);
    $$(".filltab").forEach(function (tab) {
      tab.addEventListener("click", function () {
        $$(".filltab").forEach(function (t) {
          t.classList.toggle("is-active", t === tab);
        });
        renderFillTotalBox();
        runFillPreview();
      });
    });
    $("#fillInput").addEventListener("input", function () {
      if (state.fillTimer) clearTimeout(state.fillTimer);
      state.fillTimer = setTimeout(runFillPreview, 260);
    });
    $("#fillUseTotal").addEventListener("change", function () {
      renderFillTotalBox();
      runFillPreview();
    });
    $("#fillTotal").addEventListener("input", function () {
      if (state.fillTimer) clearTimeout(state.fillTimer);
      state.fillTimer = setTimeout(runFillPreview, 300);
    });
    $("#btnPresetEdit").addEventListener("click", function () { openPresetModal("fee"); });
    $("#btnPeriodEdit").addEventListener("click", function () { openPresetModal("period"); });

    // --- 兼职时段快捷选择 ---
    $("#periodChips").addEventListener("click", function (e) {
      var chip = e.target.closest("[data-period]");
      if (chip) setPeriod(chip.dataset.period);
    });
    var periodCustom = $("#periodCustom");
    periodCustom.addEventListener("input", function () {
      // 边打字边生效，但刷新列表要防抖
      state.settings.period = periodCustom.value.trim();
      if (state.periodTimer) clearTimeout(state.periodTimer);
      state.periodTimer = setTimeout(function () {
        renderPeriod();
        scheduleSave();
      }, 300);
    });
    periodCustom.addEventListener("change", function () {
      setPeriod(periodCustom.value);
    });

    // --- 预设编辑弹窗（补贴费 / 兼职时段共用）---
    var presetModalInput = $("#presetInputModal");
    function commitPresetModal() {
      if (state.modalPresetKind === "period") {
        state.periodPresets = parsePeriodPresets(presetModalInput.value);
        renderPeriod();
      } else {
        state.feePresets = parsePresets(presetModalInput.value);
      }
      renderPresets();
      scheduleSave();
    }
    presetModalInput.addEventListener("input", function () {
      if (state.presetTimer) clearTimeout(state.presetTimer);
      state.presetTimer = setTimeout(commitPresetModal, 250);
    });
    presetModalInput.addEventListener("blur", commitPresetModal);
    $("#btnPresetReset").addEventListener("click", function () {
      if (state.modalPresetKind === "period") {
        state.periodPresets = buildPeriodPresets(state.periodStart);
        presetModalInput.value = state.periodPresets.join(", ");
        renderPeriod();
        toast("已恢复默认的月份预设", "ok");
      } else {
        state.feePresets = DEFAULT_FEE_PRESETS.slice();
        presetModalInput.value = state.feePresets.join(", ");
        toast("已恢复默认预设", "ok");
      }
      renderPresets();
      scheduleSave();
    });

    // --- 第2步：批量 / 选择 ---
    $$("[data-batch-apply]").forEach(function (btn) {
      btn.addEventListener("click", function () { applyBatch(btn.dataset.batchApply); });
    });
    $$("[data-select]").forEach(function (btn) {
      btn.addEventListener("click", function () { selectAll(btn.dataset.select); });
    });
    $("#checkAll").addEventListener("change", function (e) {
      selectAll(e.target.checked ? "all" : "none");
    });
    $("#search").addEventListener("input", function (e) {
      state.search = e.target.value;
      renderTable();
    });
    $("#btnAdd").addEventListener("click", function () { addStudent(); });
    $("#btnDeleteUnchecked").addEventListener("click", function () {
      var n = state.students.filter(function (s) { return !s.checked; }).length;
      if (!n) return toast("没有未勾选的学生", "warn");
      if (!window.confirm("删除 " + n + " 位未勾选的学生？")) return;
      state.students = state.students.filter(function (s) { return s.checked; });
      rowCache.clear();
      $("#tbody").innerHTML = "";
      renderTable(); saveNow();
    });

    function addStudent() {
      var s = newStudent({ checked: true });
      state.students.push(s);
      if (state.search) { state.search = ""; $("#search").value = ""; }
      renderTable(); scheduleSave();
      var tr = rowCache.get(s.id);
      if (tr) {
        var first = $('[data-f="studentId"]', tr);
        if (first) first.focus();
      }
    }

    // --- 表格：输入 ---
    $("#tbody").addEventListener("input", function (event) {
      var input = event.target;
      if (!input.dataset || !input.dataset.f) return;
      var tr = input.closest("tr");
      var student = studentById(tr.dataset.id);
      if (!student) return;

      var field = input.dataset.f;
      student[field] = input.value;

      // 列宽有限，长内容会被输入框截断——title 要跟着内容实时更新，
      // 否则改完名字鼠标停上去还显示旧值。
      if (field !== "checked") input.title = input.value;

      if (field === "amount") {
        student.manual = input.value.trim() !== "";
        input.classList.toggle("manual", student.manual);
        input.placeholder = student.manual ? "" : "自动";
      } else if (field === "college") {
        // 一旦手动改过，就不再算"默认填的"
        student.collegeAuto = false;
        input.classList.remove("is-default");
        input.title = input.value;
      } else if (field === "rate" || field === "hours") {
        if (!student.manual) {
          var amountInput = $('[data-f="amount"]', tr);
          if (amountInput && document.activeElement !== amountInput) {
            amountInput.value = effAmount(student);
          }
        }
      }
      renderSummary(); renderStepInfo(); renderRecap();
      scheduleSave();
    });

    // --- 表格：勾选 ---
    $("#tbody").addEventListener("change", function (event) {
      var input = event.target;
      if (!input.dataset || input.dataset.f !== "checked") return;
      var tr = input.closest("tr");
      var student = studentById(tr.dataset.id);
      if (!student) return;
      student.checked = input.checked;
      tr.dataset.checked = String(input.checked);
      tr.classList.toggle("is-unchecked", !input.checked);
      renderSummary(); renderStepInfo(); renderRecap();
      scheduleSave();
    });

    // --- 表格：删除 ---
    $("#tbody").addEventListener("click", function (event) {
      var btn = event.target.closest(".row-del");
      if (!btn) return;
      var tr = btn.closest("tr");
      var id = tr.dataset.id;
      var student = studentById(id);
      if (!student) return;
      if (!window.confirm("删除学生「" + (student.name || student.studentId || "未命名") + "」？")) return;
      state.students = state.students.filter(function (s) { return s.id !== id; });
      rowCache.delete(id);
      tr.remove();
      renderTable(); saveNow();
    });

    // --- 表格：粘贴一列 / 回车跳行 ---
    $("#tbody").addEventListener("paste", function (event) {
      var input = event.target;
      if (!input.dataset || !input.dataset.f || input.dataset.f === "checked") return;
      handlePaste(event, input);
    });

    $("#tbody").addEventListener("keydown", function (event) {
      var input = event.target;
      if (!input.dataset || !input.dataset.f || input.dataset.f === "checked") return;
      var key = event.key;

      if (key === "Enter" || key === "ArrowDown") {
        event.preventDefault();
        if (moveFocus(input, 1)) renderSummary();
      } else if (key === "ArrowUp") {
        event.preventDefault();
        moveFocus(input, -1);
      } else if (key === "ArrowRight" && input.selectionStart === input.value.length) {
        var tr = input.closest("tr");
        var cells = $$("input[data-f]", tr);
        var idx = cells.indexOf(input);
        if (idx >= 0 && idx < cells.length - 1) { event.preventDefault(); cells[idx + 1].focus(); }
      } else if (key === "ArrowLeft" && input.selectionStart === 0) {
        var tr2 = input.closest("tr");
        var cells2 = $$("input[data-f]", tr2);
        var idx2 = cells2.indexOf(input);
        if (idx2 > 0) { event.preventDefault(); cells2[idx2 - 1].focus(); }
      }
    });

    $("#tbody").addEventListener("focusin", function (event) {
      if (event.target.dataset && event.target.dataset.f) markPasteColumn(event.target);
    });
    $("#tbody").addEventListener("focusout", function () {
      setTimeout(function () {
        var active = document.activeElement;
        if (!active || active.tagName !== "INPUT" || !active.dataset.f) clearPasteColumn();
      }, 60);
    });

    // --- 第3步：导出 ---
    $("#btnExportDocx").addEventListener("click", function () { exportFiles(["docx"]); });
    $("#btnExportXls").addEventListener("click", function () { exportFiles(["xls"]); });
    $("#btnExportBoth").addEventListener("click", function () { exportFiles(["docx", "xls"]); });

    // --- 抽屉表单 ---
    $$("[data-setting]").forEach(function (el) {
      el.addEventListener("input", function () {
        state.settings[el.dataset.setting] = el.value;
        if (el.dataset.setting === "period") renderPeriod();
        scheduleSave();
      });
    });
    $$("[data-batch]").forEach(function (el) {
      el.addEventListener("input", function () {
        state.batch[el.dataset.batch] = el.value;
        renderSettings();
        scheduleSave();
      });
    });
    $("#btnTplApply").addEventListener("click", applyTemplate);
    $("#btnTplSave").addEventListener("click", function () { saveTemplate(""); });
    $("#btnTplOverwrite").addEventListener("click", function () {
      var id = selectedTemplate();
      if (!id) return toast("请先选择要覆盖的模板", "warn");
      saveTemplate(id);
    });
    $("#btnTplDelete").addEventListener("click", deleteTemplate);
    $("#tplSelect").addEventListener("change", function () { state.tplId = this.value; });
    $("#btnUploadDocx").addEventListener("click", function () { $("#fileDocx").click(); });
    $("#fileDocx").addEventListener("change", function (e) {
      if (e.target.files[0]) uploadDocx(e.target.files[0]);
    });

    // --- 弹窗 ---
    $$("[data-close]").forEach(function (btn) {
      btn.addEventListener("click", function () { btn.closest(".modal").hidden = true; });
    });
    $$(".modal").forEach(function (modal) {
      modal.addEventListener("mousedown", function (e) {
        if (e.target === modal) modal.hidden = true;
      });
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") {
        $$(".modal").forEach(function (m) { m.hidden = true; });
        closeDrawer();
      }
      if (e.ctrlKey && e.key === "s") { e.preventDefault(); saveNow(); }
    });

    window.addEventListener("beforeunload", function () {
      if (state.saveTimer) {
        try {
          navigator.sendBeacon("/api/settings",
            new Blob([JSON.stringify(payload())], { type: "application/json" }));
        } catch (err) { /* 忽略 */ }
      }
    });
  }

  // ============================================================ 深链接
  /** 支持 #step=3 / #drawer=1 / #batch=1，方便收藏某个步骤直接打开。 */
  function readHash() {
    var hash = String(window.location.hash || "").replace(/^#/, "");
    if (!hash) return {};
    var out = {};
    hash.split("&").forEach(function (part) {
      var kv = part.split("=");
      if (kv.length === 2) out[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1]);
    });
    return out;
  }

  function applyHash() {
    var opts = readHash();
    var step = Number(opts.step);
    if (step >= 1 && step <= 3) {
      if (step > 1 && !state.students.length) {
        toast("链接指向第 " + step + " 步，但还没有名单，先导入学生", "warn");
      } else {
        gotoStep(step, { keepScroll: true, silent: true });
      }
    }
    if (opts.drawer === "1") openDrawer();
    if (opts.batch === "1" && checkedStudents().length) {
      $("#fillPanel").hidden = false;
      $("#btnBatchOpen").textContent = "收起批量填入";
      renderFillTotalBox();
      runFillPreview();
    }
  }

  // ============================================================ 高度自适应
  /** 让表格填满"表格顶部到视口底部"的剩余空间，一屏能多看几行。
   *
   * 单靠 CSS 做不到：100vh 是视口高度，但表格上方还有顶栏、步骤条、
   * 金额条、工具栏，高度不固定（批量面板一展开又变了），100vh 减不掉这些。
   *
   * 量法：从 .panel 的顶端到 .table-wrap 的顶端，这段高度只取决于上面的
   * 内容，和表格自身高度无关 —— 所以不会出现"改高度 -> top 变 -> 又改
   * 高度"的正反馈（一开始就是踩了这个坑，越算越小）。 */
  function fitTableHeight() {
    var wrap = $(".table-wrap");
    if (!wrap) return;
    var panel2 = $('[data-panel="2"]');
    if (panel2 && panel2.hidden) return;
    var anchor = wrap.closest(".panel") || document.body;

    var above = wrap.getBoundingClientRect().top - anchor.getBoundingClientRect().top;
    var panelTop = anchor.getBoundingClientRect().top + window.pageYOffset;
    var avail = window.innerHeight - (panelTop + above) - 14;
    wrap.style.maxHeight = Math.max(240, Math.round(avail)) + "px";
  }

  /** 内容变化后重算（renderTable 每次都会调）。 */
  function refreshTableHeight() {
    fitTableHeight();
  }

  function bindHeightFit() {
    var queued = false;
    var run = function () {
      if (queued) return;
      queued = true;
      requestAnimationFrame(function () {
        queued = false;
        fitTableHeight();
      });
    };
    window.addEventListener("resize", run);
    window.addEventListener("orientationchange", run);
    document.addEventListener("toggle", run, true);   // <details> 展开/收起
    if (window.ResizeObserver) {
      var observed = 0;
      var ro = new ResizeObserver(run);
      $$("[data-panel]").forEach(function (p) { ro.observe(p); observed++; });
      var fill = $("#fillPanel");
      if (fill && observed) ro.observe(fill);
    }
  }

  // ============================================================ 启动
  function renderAll() {
    renderSettings();
    renderPeriod();
    renderPresets();
    renderTemplates();
    renderTable();
    renderRecap();
  }

  function boot() {
    bind();
    bindHeightFit();
    setSaveState("加载中…");
    Promise.all([api("/api/settings"), api("/api/templates")])
      .then(function (res) {
        var ws = res[0] || {};
        state.settings = Object.assign({}, state.settings, ws.settings || {});
        state.batch = Object.assign({}, state.batch, ws.batch || {});
        if (Object.prototype.toString.call(ws.feePresets) === "[object Array]") {
          state.feePresets = ws.feePresets.map(function (v) { return fmt(v); });
        }
        if (Object.prototype.toString.call(ws.periodPresets) === "[object Array]" &&
            ws.periodPresets.length) {
          state.periodPresets = ws.periodPresets.map(function (v) { return String(v); });
        }
        // 月份基准以后端为准（store.DEFAULT_PERIOD_START），前端只做兜底
        if (ws.periodStart) {
          state.periodStart = String(ws.periodStart);
          if (!state.periodPresets.length) {
            state.periodPresets = buildPeriodPresets(state.periodStart);
          }
        }
        state.students = (ws.students || []).map(normalizeStudent);
        state.templates = (res[1] && res[1].templates) || [];
        renderAll();

        var opts = readHash();
        var hasStep = Number(opts.step) >= 1 && Number(opts.step) <= 3;
        if (!hasStep) {
          // 回到上次停的那一步；没有记录就按有没有名单决定
          var saved = Number(ws.step);
          var start = (saved >= 1 && saved <= 3) ? saved : (state.students.length ? 2 : 1);
          if (start > 1 && !state.students.length) start = 1;
          gotoStep(start, { keepScroll: true, silent: true });
        }
        applyHash();
        setSaveState("已载入本地数据", "ok");
        scheduleSave();
      })
      .catch(function (err) {
        setSaveState("加载失败", "err");
        toast("无法连接本地服务：" + err.message + "（请确认启动窗口仍在运行）", "err");
      });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
