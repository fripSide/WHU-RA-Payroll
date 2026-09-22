/* 长期人员库：身份组（互斥）+ 项目组（可重叠）。数据与 revision 由 app.js 负责。 */
(function () {
  "use strict";
  function $(id) { return document.getElementById(id); }
  function esc(value) { return String(value == null ? "" : value).replace(/[&<>"']/g, function (c) {
    return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];
  }); }
  var IDENTITY = "identity", PROJECT = "project", UNGROUPED = "__none__";
  window.PayrollLibrary = function (options) {
    var identity = "", project = "", search = "", editing = null, busy = false;
    function data() { return options.getLibrary(); }
    function groups(kind) {
      return data().groups.filter(function (g) { return (g.kind || IDENTITY) === kind; });
    }
    function inGroup(person, gid) {
      if (gid === UNGROUPED) return !person.identityId;
      return !gid || person.identityId === gid;
    }
    function visible() {
      return data().people.filter(function (p) {
        return inGroup(p, identity) &&
          (!project || (p.projectIds || []).indexOf(project) >= 0) &&
          (p.name + " " + p.studentId + " " + p.college).toLowerCase().indexOf(search) >= 0;
      });
    }
    function chip(gid, name, count, active) {
      return '<button class="group-chip' + (active ? " is-active" : "") + '" type="button" data-group="' + esc(gid) +
        '">' + esc(name) + ' <span>' + count + "</span></button>";
    }
    function render() {
      var lib = data(), people = lib.people;
      if (identity && identity !== UNGROUPED && !groups(IDENTITY).some(function (g) { return g.id === identity; })) identity = "";
      if (project && !groups(PROJECT).some(function (g) { return g.id === project; })) project = "";
      $("libraryCount").textContent = people.length + " 人 · " +
        groups(IDENTITY).length + " 个身份组 · " + groups(PROJECT).length + " 个项目组";

      $("identityChips").innerHTML = [chip("", "全部", people.length, !identity)]
        .concat(groups(IDENTITY).map(function (g) {
          var count = people.filter(function (p) { return p.identityId === g.id; }).length;
          return chip(g.id, g.name, count, identity === g.id);
        }))
        .concat([chip(UNGROUPED, "未分组", people.filter(function (p) { return !p.identityId; }).length,
                      identity === UNGROUPED)])
        .join("");
      $("projectChips").innerHTML = [chip("", "全部项目", people.length, !project)]
        .concat(groups(PROJECT).map(function (g) {
          var count = people.filter(function (p) { return (p.projectIds || []).indexOf(g.id) >= 0; }).length;
          return chip(g.id, g.name, count, project === g.id);
        }))
        .join("");
      $("projectChips").hidden = groups(PROJECT).length === 0;
      $("btnRenameIdentity").hidden = $("btnDeleteIdentity").hidden = !identity || identity === UNGROUPED;
      $("btnRenameProject").hidden = $("btnDeleteProject").hidden = !project;

      var list = visible(), current = options.currentIds();
      $("btnAddGroupToBatch").disabled = !list.length || busy;
      $("libraryEmpty").hidden = list.length > 0;
      $("libraryEmpty").textContent = people.length ? "此分组或搜索条件下没有人员。" : "还没有人员，点击“新增人员”或在下面导入名单。";
      $("libraryRows").innerHTML = list.map(function (p) {
        var own = groups(IDENTITY).filter(function (g) { return g.id === p.identityId; })[0];
        var names = groups(PROJECT).filter(function (g) { return (p.projectIds || []).indexOf(g.id) >= 0; });
        return '<tr data-person="' + esc(p.id) + '"><td><strong>' + esc(p.name || "未填姓名") + '</strong><small>' + esc(p.studentId || "待补学号") + '</small></td><td>' + esc(p.college) + '</td><td>' +
          (own ? '<span class="group-tag">' + esc(own.name) + "</span>" : '<span class="tip">未分组</span>') + "</td><td>" +
          (names.length ? names.map(function (g) { return '<span class="group-tag project">' + esc(g.name) + "</span>"; }).join("") : '<span class="tip">—</span>') +
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
      editing = person ? Object.assign({}, person) : {id: options.uid()};
      var form = $("personForm"), defaults = options.defaults();
      form.reset();
      $("personTitle").textContent = person ? "编辑人员" : "新增人员";
      ["name", "studentId", "college", "rate", "hours", "reason"].forEach(function (field) {
        form.elements[field].value = editing[field] == null ? (defaults[field] || "") : editing[field];
      });
      form.elements.amount.value = editing.manual ? editing.amount : "";
      form.elements.addToBatch.checked = !person || options.currentIds().indexOf(editing.id) >= 0;

      var own = editing.identityId || (identity && identity !== UNGROUPED ? identity : "");
      $("personIdentity").innerHTML = '<option value="">（未分组）</option>' + groups(IDENTITY).map(function (g) {
        return '<option value="' + esc(g.id) + '"' + (g.id === own ? " selected" : "") + ">" + esc(g.name) + "</option>";
      }).join("");
      var chosen = editing.projectIds || [];
      $("personProjects").innerHTML = groups(PROJECT).map(function (g) {
        return '<label class="inline-check"><input type="checkbox" name="projectIds" value="' + esc(g.id) + '" ' +
          (chosen.indexOf(g.id) >= 0 ? "checked" : "") + ">" + esc(g.name) + "</label>";
      }).join("") || '<span class="tip">还没有项目组，可在上方添加</span>';
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
      person.identityId = form.elements.identityId.value;
      person.projectIds = Array.from(form.querySelectorAll('[name="projectIds"]:checked')).map(function (el) { return el.value; });
      var add = form.elements.addToBatch.checked;
      Array.from(form.querySelectorAll("button")).forEach(function (b) { b.disabled = true; });
      perform("/api/people/save", {person: person}).then(function () {
        // 学号相同可能复用了已有档案，取回真正落库的那一条。
        var saved = data().people.find(function (p) { return p.id === person.id || (person.studentId && p.studentId === person.studentId); });
        return add && saved ? options.add([saved]) : null;
      }).then(function () {
        options.toast("人员已保存，可随时继续编辑", "ok");
        if (next) open(null); else $("personModal").hidden = true;
      }).catch(function (err) { $("personError").textContent = err.message; }).then(function () {
        Array.from(form.querySelectorAll("button")).forEach(function (b) { b.disabled = false; });
      });
    }
    function pickGroup(kind) {
      var current = kind === IDENTITY ? identity : project;
      var group = data().groups.filter(function (g) { return g.id === current; })[0];
      var name = group && window.prompt(kind === IDENTITY ? "身份组新名称" : "项目组新名称", group.name);
      if (name) perform("/api/groups/save", {id: current, name: name, kind: kind}).catch(function () {});
    }
    function dropGroup(kind) {
      var current = kind === IDENTITY ? identity : project;
      var group = data().groups.filter(function (g) { return g.id === current; })[0];
      if (!group) return;
      var note = kind === IDENTITY
        ? "删除身份组「" + group.name + "」？组内人员会保留（变成未分组），档案不会被删除。"
        : "删除项目组「" + group.name + "」？人员档案保留，只是不再属于这个项目组。";
      if (window.confirm(note)) perform("/api/groups/delete", {id: current}).catch(function () {});
    }

    $("btnNewPerson").onclick = function () { open(null); };
    $("personForm").onsubmit = function (e) { e.preventDefault(); save(false); };
    $("btnSaveAndNext").onclick = function () { save(true); };
    $("librarySearch").oninput = function (e) { search = e.target.value.trim().toLowerCase(); render(); };
    $("identityChips").onclick = function (e) { var b = e.target.closest("[data-group]"); if (b) { identity = b.dataset.group; render(); } };
    $("projectChips").onclick = function (e) { var b = e.target.closest("[data-group]"); if (b) { project = b.dataset.group; render(); } };
    $("btnNewGroup").onclick = function () {
      var name = $("newGroupName").value.trim();
      var kind = $("newGroupKind").value === PROJECT ? PROJECT : IDENTITY;
      if (!name) return options.toast("请填写分组名称", "warn");
      perform("/api/groups/save", {name: name, kind: kind}).then(function () {
        $("newGroupName").value = "";
        var created = data().groups.filter(function (g) { return g.name === name; })[0];
        if (created) { if (kind === IDENTITY) identity = created.id; else project = created.id; }
        render();
      }).catch(function () {});
    };
    $("btnRenameIdentity").onclick = function () { pickGroup(IDENTITY); };
    $("btnRenameProject").onclick = function () { pickGroup(PROJECT); };
    $("btnDeleteIdentity").onclick = function () { dropGroup(IDENTITY); };
    $("btnDeleteProject").onclick = function () { dropGroup(PROJECT); };
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
    return {render:render, open:open, groups:groups};
  };
})();
