// SendShield task pane: tick each external recipient (and the attachments) before sending.
(function () {
    "use strict";

    var Core = window.SendShieldCore;
    var RT = window.SendShieldRuntime;
    var item = null;
    var config = null;
    var state = { externals: [], attachments: [], confirmed: [], attConfirmed: [] };

    function $(id) { return document.getElementById(id); }

    function show(id, visible) { $(id).hidden = !visible; }

    function setStatus(text, isError) {
        var el = $("status");
        el.textContent = text || "";
        el.className = "status" + (isError ? " error" : "");
        el.hidden = !text;
    }

    function render(recipients) {
        var cls = Core.classifyRecipients(recipients, config);
        var byAddress = {};
        for (var i = 0; i < recipients.length; i++) {
            var key = String(recipients[i].emailAddress || "").toLowerCase();
            if (!byAddress[key]) byAddress[key] = recipients[i];
        }

        state.externals = cls.externals;
        show("loading", false);
        $("internal-summary").textContent = cls.internalCount
            ? cls.internalCount + " internal recipient" + (cls.internalCount === 1 ? "" : "s")
            : "";

        if (!config.enabled || cls.externals.length === 0) {
            show("externals", false);
            show("attachments", false);
            show("actions", false);
            show("none", true);
            setStatus("");
            return;
        }
        show("none", false);
        show("externals", true);
        show("actions", true);
        $("intro").textContent = config.messages.taskpaneIntro;
        $("ext-count").textContent = "(" + cls.externals.length + ")";

        var list = $("recipient-list");
        list.innerHTML = "";
        for (var r = 0; r < cls.externals.length; r++) {
            var address = cls.externals[r];
            var info = byAddress[address] || {};
            var li = document.createElement("li");
            var label = document.createElement("label");
            var box = document.createElement("input");
            box.type = "checkbox";
            box.className = "recipient-box";
            box.value = address;
            box.checked = Core.pending([address], state.confirmed).length === 0;
            box.addEventListener("change", updateButton);

            var text = document.createElement("span");
            var addr = document.createElement("span");
            addr.className = "addr";
            addr.textContent = address;
            var chip = document.createElement("span");
            chip.className = "chip";
            chip.textContent = info.field || "";
            text.appendChild(addr);
            if (info.field) text.appendChild(chip);
            if (info.displayName && String(info.displayName).toLowerCase() !== address) {
                var name = document.createElement("span");
                name.className = "name";
                name.textContent = info.displayName;
                text.appendChild(name);
            }
            label.appendChild(box);
            label.appendChild(text);
            li.appendChild(label);
            list.appendChild(li);
        }

        renderAttachments();
        updateButton();
    }

    function renderAttachments() {
        var names = config.confirmAttachments ? state.attachments : [];
        if (!names.length) { show("attachments", false); return; }
        show("attachments", true);
        $("att-count").textContent = "(" + names.length + ")";
        var list = $("attachment-list");
        list.innerHTML = "";
        for (var i = 0; i < names.length; i++) {
            var li = document.createElement("li");
            li.textContent = names[i];
            list.appendChild(li);
        }
        $("attachments-ok").checked = Core.pending(names, state.attConfirmed).length === 0;
    }

    function recipientBoxes() {
        return Array.prototype.slice.call(document.querySelectorAll(".recipient-box"));
    }

    function attachmentsNeeded() {
        return config.confirmAttachments && state.attachments.length > 0;
    }

    function updateButton() {
        var boxes = recipientBoxes();
        var ticked = boxes.filter(function (b) { return b.checked; }).length;
        var attOk = !attachmentsNeeded() || $("attachments-ok").checked;
        $("select-all").checked = boxes.length > 0 && ticked === boxes.length;
        $("confirm").disabled = !(ticked === boxes.length && attOk);
        var left = boxes.length - ticked;
        $("progress").textContent = left > 0
            ? left + " recipient" + (left === 1 ? "" : "s") + " still to tick"
            : (attOk ? "" : "Please confirm the attachments");
        setStatus("");
    }

    function confirm() {
        var confirmedNow = recipientBoxes().map(function (b) { return b.value; });
        var confirmed = Core.union(state.confirmed, confirmedNow);
        var attConfirmed = attachmentsNeeded() ? Core.union(state.attConfirmed, state.attachments) : state.attConfirmed;
        $("confirm").disabled = true;
        RT.setSession(item, {
            ssConfirmed: JSON.stringify(confirmed),
            ssAttConfirmed: JSON.stringify(attConfirmed),
            ssBlocks: 0
        }, function (ok) {
            if (!ok) {
                setStatus("Couldn't save your confirmation. You can still send - Outlook will ask you to confirm in a dialog.", true);
                updateButton();
                return;
            }
            state.confirmed = confirmed;
            state.attConfirmed = attConfirmed;
            setStatus("Confirmed. You can now click Send.");
            $("progress").textContent = "";
            RT.sendEvent(config, "taskpane_confirmed", {
                externalCount: state.externals.length,
                externalDomains: unique(state.externals.map(Core.domainOf)),
                attachmentCount: state.attachments.length
            });
        });
    }

    function unique(list) {
        return Core.union([], list);
    }

    function reload() {
        RT.getSession(item, function (session) {
            state.confirmed = Core.parseList(session.ssConfirmed);
            state.attConfirmed = Core.parseList(session.ssAttConfirmed);
            RT.getAttachmentNames(item, function (names) {
                state.attachments = names;
                RT.getRecipients(item, function (list) { render(list); });
            });
        });
    }

    Office.onReady(function () {
        item = Office.context.mailbox.item;
        RT.loadConfig(function (cfg) {
            config = cfg;
            show("legacy", !RT.sessionAvailable(item));
            reload();
            // Pick up the latest settings, then re-render if they changed.
            RT.refreshConfig(true, function (changed) {
                if (changed) RT.loadConfig(function (fresh) { config = fresh; reload(); });
            });
        });

        $("confirm").addEventListener("click", confirm);
        $("attachments-ok").addEventListener("change", updateButton);
        $("select-all").addEventListener("change", function (e) {
            recipientBoxes().forEach(function (b) { b.checked = e.target.checked; });
            updateButton();
        });

        try { item.addHandlerAsync(Office.EventType.RecipientsChanged, reload); } catch (e) { /* older hosts */ }
        try { item.addHandlerAsync(Office.EventType.AttachmentsChanged, reload); } catch (e) { /* older hosts */ }
    });
})();
