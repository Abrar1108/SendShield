// SendShield task pane: confirm external recipients (and attachments) before sending.
// Every visible string comes from config.messages, which admins edit in the dashboard.
(function () {
    "use strict";

    var Core = window.SendShieldCore;
    var RT = window.SendShieldRuntime;
    var item = null;
    var config = null;
    var state = { externals: [], attachments: [], confirmed: [], attConfirmed: [] };

    function $(id) { return document.getElementById(id); }
    function show(id, visible) { $(id).hidden = !visible; }
    function msg(key, values) { return Core.fillTemplate(config.messages[key], values || {}); }

    function setStatus(text, isError) {
        var el = $("status");
        el.textContent = text || "";
        el.className = "status" + (isError ? " error" : "");
        el.hidden = !text;
    }

    function eachStyle() { return config.confirmStyle !== "all"; }

    function applyStaticText() {
        document.title = config.messages.taskpaneTitle;
        $("title").textContent = msg("taskpaneTitle");
        $("intro").textContent = msg("taskpaneIntro");
        $("recipients-heading").textContent = msg("recipientsHeading");
        $("select-all-label").textContent = msg("selectAll");
        $("attachments-heading").textContent = msg("attachmentsHeading");
        $("attachments-label").textContent = msg("attachmentsCheckbox");
        $("confirm").textContent = msg("confirmButton");
        $("none").textContent = msg("noExternal");
        $("legacy").textContent = msg("legacyNotice");
    }

    function showOnly(message) {
        show("loading", false);
        show("externals", false);
        show("attachments", false);
        show("actions", false);
        $("none").textContent = message;
        show("none", true);
        setStatus("");
    }

    function render(recipients) {
        applyStaticText();
        var cls = Core.classifyRecipients(recipients, config);
        var byAddress = {};
        for (var i = 0; i < recipients.length; i++) {
            var key = String(recipients[i].emailAddress || "").toLowerCase();
            if (!byAddress[key]) byAddress[key] = recipients[i];
        }
        state.externals = cls.externals;
        $("internal-summary").textContent = cls.internalCount ? msg("internalSummary", { count: cls.internalCount }) : "";

        if (!Core.isPilotUser(RT.userEmail(), config)) { showOnly(msg("pilotInactive")); return; }
        if (!config.enabled || cls.externals.length === 0) { showOnly(msg("noExternal")); return; }

        show("loading", false);
        show("none", false);
        show("externals", true);
        show("actions", true);
        $("ext-count").textContent = "(" + cls.externals.length + ")";
        show("select-all-row", eachStyle());
        show("confirm-all-row", !eachStyle());
        $("confirm-all-label").textContent = msg("confirmAllCheckbox", { count: cls.externals.length });
        $("confirm-all").checked = Core.pending(cls.externals, state.confirmed).length === 0;

        var list = $("recipient-list");
        list.className = "list" + (eachStyle() ? "" : " plain");
        list.innerHTML = "";
        for (var r = 0; r < cls.externals.length; r++) {
            var address = cls.externals[r];
            var info = byAddress[address] || {};
            var li = document.createElement("li");
            var holder = document.createElement(eachStyle() ? "label" : "div");
            holder.className = "recipient";
            if (eachStyle()) {
                var box = document.createElement("input");
                box.type = "checkbox";
                box.className = "recipient-box";
                box.value = address;
                box.checked = Core.pending([address], state.confirmed).length === 0;
                box.addEventListener("change", updateButton);
                holder.appendChild(box);
            }
            var text = document.createElement("span");
            var addr = document.createElement("span");
            addr.className = "addr";
            addr.textContent = address;
            text.appendChild(addr);
            if (info.field) {
                var chip = document.createElement("span");
                chip.className = "chip";
                chip.textContent = info.field;
                text.appendChild(chip);
            }
            if (info.displayName && String(info.displayName).toLowerCase() !== address) {
                var name = document.createElement("span");
                name.className = "name";
                name.textContent = info.displayName;
                text.appendChild(name);
            }
            holder.appendChild(text);
            li.appendChild(holder);
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
        var left;
        if (eachStyle()) {
            var boxes = recipientBoxes();
            var ticked = boxes.filter(function (b) { return b.checked; }).length;
            $("select-all").checked = boxes.length > 0 && ticked === boxes.length;
            left = boxes.length - ticked;
        } else {
            left = $("confirm-all").checked ? 0 : state.externals.length;
        }
        var attOk = !attachmentsNeeded() || $("attachments-ok").checked;
        $("confirm").disabled = !(left === 0 && attOk);
        $("progress").textContent = left > 0
            ? (eachStyle() ? msg("remainingRecipients", { count: left }) : "")
            : (attOk ? "" : msg("remainingAttachments"));
        setStatus("");
    }

    function confirm() {
        var confirmed = Core.union(state.confirmed, state.externals);
        var attConfirmed = attachmentsNeeded() ? Core.union(state.attConfirmed, state.attachments) : state.attConfirmed;
        $("confirm").disabled = true;
        RT.setSession(item, {
            ssConfirmed: JSON.stringify(confirmed),
            ssAttConfirmed: JSON.stringify(attConfirmed),
            ssBlocks: 0
        }, function (ok) {
            if (!ok) {
                setStatus(msg("saveFailed"), true);
                updateButton();
                return;
            }
            state.confirmed = confirmed;
            state.attConfirmed = attConfirmed;
            setStatus(msg("confirmedMessage"));
            $("progress").textContent = "";
            RT.sendEvent(config, "taskpane_confirmed", {
                externalCount: state.externals.length,
                externalDomains: Core.union([], state.externals.map(Core.domainOf)),
                attachmentCount: state.attachments.length
            });
        });
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
        $("confirm-all").addEventListener("change", updateButton);
        $("select-all").addEventListener("change", function (e) {
            recipientBoxes().forEach(function (b) { b.checked = e.target.checked; });
            updateButton();
        });

        try { item.addHandlerAsync(Office.EventType.RecipientsChanged, reload); } catch (e) { /* older hosts */ }
        try { item.addHandlerAsync(Office.EventType.AttachmentsChanged, reload); } catch (e) { /* older hosts */ }
    });
})();
