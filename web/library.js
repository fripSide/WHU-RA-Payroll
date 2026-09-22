/* Long-lived people and arbitrary groups. The app owns persistence and revisions. */
(function () {
  "use strict";
  function $(id) { return document.getElementById(id); }
  function esc(value) { return String(value == null ? "" : value).replace(/[&<>"']/g, function (c) {
    return {"&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"}[c];
  }); }
  window.PayrollLibrary = function (options) {
    var group = "", search = "", editing = null, busy = false;
    function data() { return options.getLibrary(); }
    function visible() {
      return data().people.filter(function (p) {
        return (!group || (group === "ungrouped" ? !p.groupIds.length : p.groupIds.indexOf(group) >= 0)) &&
          (p.name + " " + p.studentId + " " + p.college).toLowerCase().indexOf(search) >= 0;
      });
    }
    function render() {
      var lib = data();
      if (group && group !== "ungrouped" && !lib.groups.some(function (g) { return g.id === group; })) group = "";
      $("libraryCount").textContent = lib.people.length + " 人 · " + lib.groups.length + " 个分组";
      var chips = [{id:"",name:"全部人员"}, {id:"ungrouped",name:"未分组"}].concat(lib.groups);
      $("groupChips").innerHTML = chips.map(function (g) {
        var count = lib.people.filter(function (p) { return !g.id || (g.id === "ungrouped" ? !p.groupIds.length : p.groupIds.indexOf(g.id) >= 0); }).length;
        return '<button class="group-chip ' + (group === g.id ? "is-active" : "") + '" type="button" data-group="' + esc(g.id) + '">' + esc(g.name) + ' <span>' + count + '</span></button>';
      }).join("");
      $("btnRenameGroup").hidden = $("btnDeleteGroup").hidden = !group || group === "ungrouped";
      var list = visible(), current = options.currentIds();
      $("btnAddGroupToBatch").disabled = !list.length || busy;
      $("libraryEmpty").hidden = list.length > 0;
      $("libraryEmpty").textContent = lib.people.length ? "此分组或搜索条件下没有人员。" : "还没有人员，点击“新增人员”或在下面导入名单。";
      $("libraryRows").innerHTML = list.map(function (p) {
        var names = lib.groups.filter(function (g) { return p.groupIds.indexOf(g.id) >= 0; }).map(function (g) { return g.name; });
        return '<tr data-person="' + esc(p.id) + '"><td><strong>' + esc(p.name || "未填姓名") + '</strong><small>' + esc(p.studentId || "待补学号") + '</small></td><td>' + esc(p.college) + '</td><td>' +
          (names.length ? names.map(function (n) { return '<span class="group-tag">' + esc(n) + '</span>'; }).join("") : '<span class="tip">未分组</span>') +
          '</td><td class="library-actions"><button class="btn btn-mini" type="button" data-action="add">' + (current.indexOf(p.id) >= 0 ? "已在本次" : "加入本次") +
          '</button><button class="btn btn-mini" type="button" data-action="edit">编辑</button><button class="btn btn-danger-ghost" type="button" data-action="delete">删除人员</button></td></tr>';
      }).join("");
    }
    function perform(path, body) {
      if (busy) return Promise.reject(new Error("请等待当前保存完成"));
      busy = true;
      return options.mutate(path, body).then(function (result) { busy = false; render(); return result; }, function (err) {
        busy = false; render(); options.toast(err.message, "err"); throw err;
      });
    }
    function open(person) {
      editing = person ? Object.assign({}, person) : {id: options.uid(), groupIds: group && group !== "ungrouped" ? [group] : []};
      var form = $("personForm"), defaults = options.defaults();
      form.reset();
      $("personTitle").textContent = person ? "编辑人员" : "新增人员";
      ["name", "studentId", "college", "rate", "hours", "reason"].forEach(function (field) {
        form.elements[field].value = editing[field] == null ? (defaults[field] || "") : editing[field];
      });
      form.elements.amount.value = editing.manual ? editing.amount : "";
      form.elements.addToBatch.checked = !person || options.currentIds().indexOf(editing.id) >= 0;
      $("personGroups").innerHTML = data().groups.map(function (g) {
        return '<label class="inline-check"><input type="checkbox" name="groupIds" value="' + esc(g.id) + '" ' +
          (editing.groupIds.indexOf(g.id) >= 0 ? "checked" : "") + '>' + esc(g.name) + '</label>';
      }).join("") || '<span class="tip">可关闭窗口后先创建分组</span>';
      $("personError").textContent = "";
      $("personModal").hidden = false;
      form.elements.name.focus();
    }
    function save(next) {
      var form = $("personForm");
      if (busy || !form.reportValidity()) return;
      var person = Object.assign({}, editing);
      ["name", "studentId", "college", "rate", "hours", "amount", "reason"].forEach(function (key) { person[key] = form.elements[key].value.trim(); });
      person.manual = person.amount !== "";
      person.collegeAuto = false;
      person.groupIds = Array.from(form.querySelectorAll('[name="groupIds"]:checked')).map(function (el) { return el.value; });
      var add = form.elements.addToBatch.checked;
      Array.from(form.querySelectorAll("button")).forEach(function (b) { b.disabled = true; });
      perform("/api/people/save", {person: person}).then(function () {
        // A matching student ID may have reused an existing persistent record.
        var saved = data().people.find(function (p) { return p.id === person.id || (person.studentId && p.studentId === person.studentId); });
        return add && saved ? options.add([saved]) : null;
      }).then(function () {
        options.toast("人员已保存，可随时继续编辑", "ok");
        if (next) open(null); else $("personModal").hidden = true;
      }).catch(function (err) { $("personError").textContent = err.message; }).then(function () {
        Array.from(form.querySelectorAll("button")).forEach(function (b) { b.disabled = false; });
      });
    }
    $("btnNewPerson").onclick = function () { open(null); };
    $("personForm").onsubmit = function (e) { e.preventDefault(); save(false); };
    $("btnSaveAndNext").onclick = function () { save(true); };
    $("librarySearch").oninput = function (e) { search = e.target.value.trim().toLowerCase(); render(); };
    $("groupChips").onclick = function (e) { var b = e.target.closest("[data-group]"); if (b) { group = b.dataset.group; render(); } };
    $("btnNewGroup").onclick = function () {
      var name = $("newGroupName").value.trim();
      if (!name) return options.toast("请填写分组名称", "warn");
      perform("/api/groups/save", {name:name}).then(function () {
        $("newGroupName").value = "";
        group = data().groups.find(function (g) { return g.name === name; }).id; render();
      }).catch(function () {});
    };
    $("btnRenameGroup").onclick = function () {
      var current = data().groups.find(function (g) { return g.id === group; });
      var name = current && window.prompt("分组新名称", current.name);
      if (name) perform("/api/groups/save", {id:group,name:name}).catch(function () {});
    };
    $("btnDeleteGroup").onclick = function () {
      if (window.confirm("删除这个分组？组内人员会保留，仍可在“全部人员”中编辑。")) {
        perform("/api/groups/delete", {id:group}).catch(function () {});
      }
    };
    $("btnAddGroupToBatch").onclick = function () { options.add(visible()).catch(function () {}); };
    $("libraryRows").onclick = function (e) {
      var button = e.target.closest("[data-action]");
      if (!button || busy) return;
      var person = data().people.find(function (p) { return p.id === button.closest("tr").dataset.person; });
      if (!person) return;
      if (button.dataset.action === "edit") open(person);
      if (button.dataset.action === "add") options.add([person]).catch(function () {});
      if (button.dataset.action === "delete" && window.confirm("从人员库删除“" + person.name + "”？该人员也会从本次名单移除。此操作只删除这一人。")) {
        perform("/api/people/delete", {id:person.id}).catch(function () {});
      }
    };
    $("btnLibraryBackup").onclick = function () {
      options.backup().then(function (backup) {
        var url = URL.createObjectURL(new Blob([JSON.stringify(backup,null,2)], {type:"application/json"}));
        var a = document.createElement("a"); a.href = url; a.download = "人员库备份_" + new Date().toISOString().slice(0,10) + ".json"; a.click();
        setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
      }).catch(function (err) { options.toast(err.message,"err"); });
    };
    $("btnLibraryRestore").onclick = function () { $("fileLibraryRestore").click(); };
    $("fileLibraryRestore").onchange = function (e) {
      var file = e.target.files[0];
      if (!file) return;
      file.text().then(function (text) {
        var backup = JSON.parse(text);
        if (!window.confirm("合并此人员库备份？相同学号的人员将更新，现有其他人员会保留。")) return;
        return perform("/api/library/import", {backup:backup});
      }).catch(function (err) { options.toast("导入失败：" + err.message,"err"); }).then(function () { e.target.value = ""; });
    };
    return {render:render, open:open};
  };
})();
