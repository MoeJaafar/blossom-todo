// Blossom To-Do — Outlook task pane, desktop app and Android app
// With a Blossom account (src/cloud.js) tasks sync through the cloud to every
// device. Without one, they're saved in Outlook roaming settings, or in
// localStorage when the page is opened outside Outlook.

(() => {
  const STORAGE_KEY = "blossomTodos";
  const $ = (id) => document.getElementById(id);
  const cloud = window.BlossomCloud?.configured ? window.BlossomCloud : null;
  const native = window.Capacitor?.isNativePlatform?.() ? window.Capacitor.Plugins.Blossom : null;

  let todos = [];
  let filter = "all";
  let inOutlook = false;
  let user = null; // signed-in Blossom account (cloud mode)
  let unsubscribe = null;

  // ---------- Storage ----------------------------------------------------

  // Local storage for when there's no account (and for importing on first sign-in)
  const store = {
    load() {
      try {
        if (inOutlook) return Office.context.roamingSettings.get(STORAGE_KEY) || [];
        return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
      } catch {
        return [];
      }
    },
    save() {
      if (inOutlook) {
        Office.context.roamingSettings.set(STORAGE_KEY, todos);
        Office.context.roamingSettings.saveAsync((res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            console.warn("Blossom To-Do: could not save", res.error);
          }
        });
      } else {
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(todos));
        } catch {}
      }
    },
    clear() {
      if (inOutlook) {
        Office.context.roamingSettings.remove(STORAGE_KEY);
        Office.context.roamingSettings.saveAsync(() => {});
      } else {
        try {
          localStorage.removeItem(STORAGE_KEY);
        } catch {}
      }
    },
  };

  // Persist one changed task (cloud: just that task; local: the whole list)
  function saved(t) {
    t.updated = Date.now();
    if (user) cloud.put(user.uid, t);
    else store.save();
  }

  function removed(ids) {
    if (user) ids.forEach((id) => cloud.remove(user.uid, id));
    else store.save();
  }

  // ---------- Tasks ------------------------------------------------------

  function addTodo(text, email) {
    text = text.trim().slice(0, 300);
    if (!text) return;
    const t = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      text,
      done: false,
      created: Date.now(),
      remindAt: null,
      ...(email ? { email } : {}),
    };
    todos.unshift(t);
    saved(t);
    render();
  }

  function toggle(id, checkEl) {
    const t = todos.find((x) => x.id === id);
    if (!t) return;
    t.done = !t.done;
    saved(t);
    if (t.done && checkEl) petals.burst(checkEl.getBoundingClientRect());
    render();
  }

  function remove(id) {
    todos = todos.filter((x) => x.id !== id);
    removed([id]);
    render();
  }

  function rename(id, text) {
    const t = todos.find((x) => x.id === id);
    text = text.trim().slice(0, 300);
    if (t && text && text !== t.text) {
      t.text = text;
      saved(t);
    }
    render();
  }

  function setReminder(id, when) {
    const t = todos.find((x) => x.id === id);
    if (!t) return;
    t.remindAt = when || null;
    saved(t);
    render();
  }

  // ---------- Rendering --------------------------------------------------

  function render() {
    syncNative();
    const list = $("list");
    list.replaceChildren();

    const visible = todos.filter((t) =>
      filter === "open" ? !t.done : filter === "done" ? t.done : true
    );

    for (const t of visible) list.appendChild(renderItem(t));

    const done = todos.filter((t) => t.done).length;
    $("progressText").textContent = `${done} of ${todos.length} bloomed`;
    $("progressFill").style.width = todos.length ? `${(done / todos.length) * 100}%` : "0";

    const empty = visible.length === 0;
    $("empty").hidden = !empty;
    $("emptyText").textContent =
      todos.length === 0
        ? "Your garden is empty. Plant a task!"
        : filter === "done"
          ? "Nothing has bloomed yet. You've got this!"
          : "Everything has bloomed! Time for tea.";
    $("clearDone").hidden = done === 0;
  }

  function renderItem(t) {
    const li = document.createElement("li");
    li.className = "item" + (t.done ? " done" : "");

    const check = document.createElement("button");
    check.type = "button";
    check.className = "check";
    check.setAttribute("role", "checkbox");
    check.setAttribute("aria-checked", String(t.done));
    check.setAttribute("aria-label", t.done ? "Mark as not done" : "Mark as done");
    check.addEventListener("click", () => toggle(t.id, check));

    const body = document.createElement("div");
    body.className = "item-body";

    const text = document.createElement("span");
    text.className = "item-text";
    text.textContent = t.text;
    text.title = "Double-click to edit";
    text.addEventListener("dblclick", () => startEdit(t, body, text));
    body.appendChild(text);

    const meta = document.createElement("div");
    meta.className = "item-meta";
    if (t.email) {
      const canOpen = inOutlook && t.email.itemId && Office.context.mailbox.displayMessageForm;
      const link = document.createElement(canOpen ? "button" : "span");
      link.className = "email-link";
      link.textContent = "✉ " + (t.email.subject || "(no subject)");
      if (canOpen) {
        link.type = "button";
        link.title = "Open this email";
        link.addEventListener("click", () => Office.context.mailbox.displayMessageForm(t.email.itemId));
      }
      meta.appendChild(link);
    } else {
      meta.append(
        new Date(t.created).toLocaleDateString(undefined, {
          month: "short",
          day: "numeric",
        })
      );
    }
    if (t.remindAt) {
      const chip = document.createElement("span");
      chip.className = "remind-chip" + (!t.done && t.remindAt < Date.now() ? " late" : "");
      chip.innerHTML = '<span class="bell-icon" aria-hidden="true"></span>';
      chip.append(formatWhen(t.remindAt));
      chip.title = "Reminder";
      meta.appendChild(chip);
    }
    body.appendChild(meta);

    const bell = document.createElement("button");
    bell.type = "button";
    bell.className = "icon-btn bell" + (t.remindAt ? " on" : "");
    bell.innerHTML = '<span class="bell-icon" aria-hidden="true"></span>';
    bell.setAttribute("aria-label", t.remindAt ? "Change reminder" : "Add reminder");
    bell.title = t.remindAt ? "Change reminder" : "Remind me";
    bell.addEventListener("click", () => openReminder(t, body));

    const del = document.createElement("button");
    del.type = "button";
    del.className = "icon-btn delete";
    del.textContent = "×";
    del.setAttribute("aria-label", "Delete task");
    del.addEventListener("click", () => remove(t.id));

    li.append(check, body, bell, del);
    return li;
  }

  // ---------- Reminders --------------------------------------------------

  function formatWhen(ms) {
    const d = new Date(ms);
    const day = (offset) => {
      const x = new Date();
      x.setDate(x.getDate() + offset);
      return x.toDateString();
    };
    const time = d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
    if (d.toDateString() === day(0)) return `Today ${time}`;
    if (d.toDateString() === day(1)) return `Tomorrow ${time}`;
    return `${d.toLocaleDateString(undefined, { month: "short", day: "numeric" })} ${time}`;
  }

  // <input type="datetime-local"> works in local time without a zone
  const toLocalInput = (ms) => {
    const d = new Date(ms);
    return new Date(ms - d.getTimezoneOffset() * 60_000).toISOString().slice(0, 16);
  };

  function quickTimes() {
    const now = new Date();
    const inHour = new Date(now.getTime() + 60 * 60_000);
    inHour.setSeconds(0, 0);
    const tonight = new Date(now);
    tonight.setHours(19, 0, 0, 0);
    const tomorrow = new Date(now);
    tomorrow.setDate(now.getDate() + 1);
    tomorrow.setHours(9, 0, 0, 0);
    const opts = [["In 1 hour", inHour]];
    if (tonight > inHour) opts.push(["Tonight", tonight]);
    opts.push(["Tomorrow 9am", tomorrow]);
    return opts;
  }

  function openReminder(t, body) {
    document.querySelector(".remind-row")?.remove();
    const row = document.createElement("div");
    row.className = "remind-row";

    const quick = document.createElement("div");
    quick.className = "remind-quick";
    for (const [label, date] of quickTimes()) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "chip-btn";
      b.textContent = label;
      b.addEventListener("click", () => setReminder(t.id, date.getTime()));
      quick.appendChild(b);
    }

    const pick = document.createElement("div");
    pick.className = "remind-pick";
    const input = document.createElement("input");
    input.type = "datetime-local";
    input.setAttribute("aria-label", "Reminder date and time");
    input.value = toLocalInput(t.remindAt || quickTimes().at(-1)[1].getTime());
    const set = document.createElement("button");
    set.type = "button";
    set.className = "btn btn-primary btn-small";
    set.textContent = "Set";
    set.addEventListener("click", () => {
      const when = new Date(input.value).getTime();
      if (when) setReminder(t.id, when);
    });
    pick.append(input, set);

    const actions = document.createElement("div");
    actions.className = "remind-actions";
    if (t.remindAt) {
      const clear = document.createElement("button");
      clear.type = "button";
      clear.className = "link-btn";
      clear.textContent = "remove reminder";
      clear.addEventListener("click", () => setReminder(t.id, null));
      actions.appendChild(clear);
    }
    const cancel = document.createElement("button");
    cancel.type = "button";
    cancel.className = "link-btn";
    cancel.textContent = "cancel";
    cancel.addEventListener("click", () => render());
    actions.appendChild(cancel);

    row.append(quick, pick, actions);
    row.addEventListener("keydown", (e) => e.key === "Escape" && render());
    body.appendChild(row);
    if (native) native.requestNotificationPermission().catch(() => {});
  }

  function startEdit(t, body, textEl) {
    const input = document.createElement("input");
    input.type = "text";
    input.className = "item-edit";
    input.maxLength = 300;
    input.value = t.text;
    textEl.replaceWith(input);
    input.focus();
    input.select();
    let finished = false;
    const finish = (save) => {
      if (finished) return;
      finished = true;
      save ? rename(t.id, input.value) : render();
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") finish(true);
      if (e.key === "Escape") finish(false);
    });
    input.addEventListener("blur", () => finish(true));
  }

  // ---------- Account & sync ---------------------------------------------

  function showView(name) {
    $("loadingView").hidden = name !== "loading";
    $("authView").hidden = name !== "auth";
    $("todoView").hidden = name !== "todos";
    $("account").hidden = !user;
    if (name !== "todos") $("clearDone").hidden = true;
    if (user) $("accountEmail").textContent = user.email;
  }

  function startData() {
    if (!cloud) {
      todos = store.load();
      showView("todos");
      render();
      return;
    }
    showView("loading");
    cloud.onUser(async (u) => {
      unsubscribe?.();
      unsubscribe = null;
      user = u;
      if (!u) {
        todos = [];
        native?.signOut().catch(() => {});
        showView("auth");
        return;
      }
      showView("todos");
      await importLocalTasks(u);
      unsubscribe = cloud.subscribe(u.uid, (list) => {
        todos = list;
        // Don't yank an open editor away when another device changes something
        if (document.querySelector(".item-edit, .remind-row")) syncNative();
        else render();
      });
    });
  }

  // Tasks saved on this device before accounts existed move into the account once
  async function importLocalTasks(u) {
    const local = store.load();
    if (!local.length) return;
    try {
      await cloud.importAll(u.uid, local);
      store.clear();
    } catch (e) {
      console.warn("Blossom To-Do: could not import local tasks", e);
    }
  }

  function wireAuthForm() {
    const msg = (text, ok) => {
      $("authMsg").textContent = text;
      $("authMsg").classList.toggle("ok", !!ok);
    };
    const run = async (action) => {
      const email = $("authEmail").value.trim();
      const password = $("authPassword").value;
      if (!email) return msg("Enter your email first.");
      if (action !== "reset" && !password) return msg("Enter your password.");
      $("authView").classList.add("busy");
      msg("");
      try {
        if (action === "signIn") await cloud.signIn(email, password);
        if (action === "signUp") await cloud.signUp(email, password);
        if (action === "reset") {
          await cloud.resetPassword(email);
          msg("Check your inbox for a reset link.", true);
        }
        $("authPassword").value = "";
      } catch (e) {
        msg(cloud.friendlyError(e));
      } finally {
        $("authView").classList.remove("busy");
      }
    };
    $("authForm").addEventListener("submit", (e) => {
      e.preventDefault();
      run("signIn");
    });
    $("signUp").addEventListener("click", () => run("signUp"));
    $("forgot").addEventListener("click", () => run("reset"));
    $("signOut").addEventListener("click", () => cloud.signOut());
  }

  // Android app: hand tasks to the home-screen widget and reminder alarms
  let nativeTimer = 0;
  function syncNative() {
    if (!native) return;
    clearTimeout(nativeTimer);
    nativeTimer = setTimeout(() => {
      const creds = user && cloud.nativeCredentials();
      native.sync({ tasks: JSON.stringify(todos), creds: creds ? JSON.stringify(creds) : "" }).catch(() => {});
    }, 300);
  }

  // Android widget "+" button opens the app ready to type
  function checkNativeAction() {
    native?.takeLaunchAction().then(({ action }) => {
      if (action === "new" && !$("todoView").hidden) $("newTask").focus();
    }).catch(() => {});
  }

  // ---------- Farm date --------------------------------------------------

  function renderDate() {
    const now = new Date();
    const seasons = ["Winter", "Spring", "Summer", "Fall"];
    $("season").textContent = seasons[Math.floor(((now.getMonth() + 1) % 12) / 3)];
    $("today").textContent = now.toLocaleDateString(undefined, {
      weekday: "short",
      month: "short",
      day: "numeric",
    });
  }

  // ---------- Current email -> task --------------------------------------

  function getCurrentEmail(cb) {
    const item = inOutlook && Office.context.mailbox.item;
    if (!item) return cb(null);
    if (typeof item.subject === "string") {
      // Read mode: subject + id so we can reopen the email later
      return cb({ subject: item.subject, itemId: item.itemId });
    }
    if (item.subject && item.subject.getAsync) {
      // Compose mode: drafts have no stable id, keep just the subject
      item.subject.getAsync((res) => cb({ subject: res.value || "", itemId: null }));
      return;
    }
    cb(null);
  }

  function refreshEmailButton() {
    const btn = $("addEmail");
    btn.hidden = !(inOutlook && Office.context.mailbox.item);
  }

  // ---------- Theme ------------------------------------------------------

  let hostTheme = null; // "default" | "dark" | "contrast" when running as an Outlook/Teams app

  function applyTheme() {
    let dark = matchMedia("(prefers-color-scheme: dark)").matches;
    if (hostTheme) dark = hostTheme !== "default";
    try {
      const bg = Office.context.officeTheme?.bodyBackgroundColor;
      if (bg) {
        const n = parseInt(bg.replace("#", ""), 16);
        const lum = ((n >> 16) & 255) * 0.299 + ((n >> 8) & 255) * 0.587 + (n & 255) * 0.114;
        dark = lum < 128;
      }
    } catch {}
    document.documentElement.dataset.theme = dark ? "dark" : "light";
  }

  // ---------- Pixel petals ----------------------------------------------

  const petals = (() => {
    const PAL = { O: "#c2408a", P: "#ff9ccb", L: "#ffd6ea" };
    const LILAC = { O: "#7a4bb0", P: "#c9a2f2", L: "#efe2ff" };
    const PETAL = [".OO.", "OLPO", "OPPO", ".OP.", "..O."];
    const SMALL = ["LP", "PO"];
    const reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;
    const canvas = $("petals");
    const ctx = canvas.getContext("2d");
    const sprites = [sprite(PETAL, PAL), sprite(SMALL, LILAC), sprite(PETAL, LILAC)];
    let flakes = [];
    let bursts = [];
    let last = 0;

    function sprite(grid, pal) {
      const c = document.createElement("canvas");
      c.width = grid[0].length;
      c.height = grid.length;
      const g = c.getContext("2d");
      grid.forEach((row, y) =>
        [...row].forEach((ch, x) => {
          if (pal[ch]) {
            g.fillStyle = pal[ch];
            g.fillRect(x, y, 1, 1);
          }
        })
      );
      return c;
    }

    function flake(initial) {
      return {
        img: sprites[Math.floor(Math.random() * sprites.length)],
        x: Math.random() * innerWidth,
        y: initial ? Math.random() * innerHeight : -16,
        vy: 12 + Math.random() * 16,
        vx: -4 + Math.random() * 12,
        sway: 8 + Math.random() * 14,
        phase: Math.random() * 6.28,
        alpha: 0.35 + Math.random() * 0.3,
      };
    }

    function resize() {
      const dpr = devicePixelRatio || 1;
      canvas.width = Math.round(innerWidth * dpr);
      canvas.height = Math.round(innerHeight * dpr);
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.imageSmoothingEnabled = false;
    }

    function draw(img, x, y, scale, squash, alpha) {
      const w = img.width * scale;
      const sw = Math.max(scale, Math.round(Math.abs(squash) * w));
      ctx.globalAlpha = alpha;
      ctx.drawImage(img, Math.round(x + (w - sw) / 2), Math.round(y), sw, img.height * scale);
    }

    function frame(t) {
      requestAnimationFrame(frame);
      const dt = Math.min(0.05, (t - (last || t)) / 1000);
      last = t;
      ctx.clearRect(0, 0, innerWidth, innerHeight);

      // gentle background drift (behind the panel)
      const target = reduce ? 0 : Math.round(innerWidth / 45);
      while (flakes.length < target) flakes.push(flake(false));
      for (let i = 0; i < flakes.length; i++) {
        const f = flakes[i];
        f.phase += dt * 1.6;
        f.y += f.vy * dt;
        f.x += (f.vx + Math.sin(f.phase) * f.sway) * dt;
        if (f.y > innerHeight + 20) flakes[i] = flake(false);
        draw(f.img, f.x, f.y, 2, Math.cos(f.phase * 1.3), f.alpha);
      }

      // completion bursts
      bursts = bursts.filter((p) => (p.life -= dt) > 0);
      for (const p of bursts) {
        p.vy += 260 * dt;
        p.x += p.vx * dt;
        p.y += p.vy * dt;
        p.phase += dt * 8;
        draw(p.img, p.x, p.y, 2, Math.cos(p.phase), Math.min(1, p.life * 2));
      }
      ctx.globalAlpha = 1;
    }

    addEventListener("resize", resize);
    resize();
    flakes = Array.from({ length: reduce ? 0 : Math.round(innerWidth / 45) }, () => flake(true));
    requestAnimationFrame(frame);

    return {
      burst(rect) {
        if (reduce) return;
        canvas.style.zIndex = 2; // draw the burst above the panel
        for (let i = 0; i < 10; i++) {
          const a = (Math.PI * 2 * i) / 10 + Math.random() * 0.4;
          const speed = 70 + Math.random() * 70;
          bursts.push({
            img: sprites[i % sprites.length],
            x: rect.left + rect.width / 2,
            y: rect.top + rect.height / 2,
            vx: Math.cos(a) * speed,
            vy: Math.sin(a) * speed - 90,
            phase: Math.random() * 6,
            life: 0.9 + Math.random() * 0.4,
          });
        }
        clearTimeout(this._t);
        this._t = setTimeout(() => (canvas.style.zIndex = 0), 1400);
      },
    };
  })();

  // ---------- Wire up ----------------------------------------------------

  function init() {
    applyTheme();
    renderDate();
    if (cloud) wireAuthForm();
    startData();
    refreshEmailButton();

    $("addForm").addEventListener("submit", (e) => {
      e.preventDefault();
      addTodo($("newTask").value);
      $("newTask").value = "";
      $("newTask").focus();
    });

    $("addEmail").addEventListener("click", () => {
      getCurrentEmail((email) => {
        if (!email) return;
        addTodo("Follow up: " + (email.subject || "(no subject)"), email);
      });
    });

    document.querySelectorAll(".tab").forEach((tab) =>
      tab.addEventListener("click", () => {
        filter = tab.dataset.filter;
        document.querySelectorAll(".tab").forEach((t) => {
          t.classList.toggle("active", t === tab);
          t.setAttribute("aria-selected", String(t === tab));
        });
        render();
      })
    );

    $("clearDone").addEventListener("click", () => {
      const doneIds = todos.filter((t) => t.done).map((t) => t.id);
      todos = todos.filter((t) => !t.done);
      removed(doneIds);
      render();
    });

    if (native) {
      checkNativeAction();
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") {
          checkNativeAction();
          render(); // refresh "Today"/"late" labels
        }
      });
    }

    // When the pane is pinned, Outlook keeps it open as you switch emails
    if (inOutlook) {
      try {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, refreshEmailButton);
      } catch {}
      try {
        Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyTheme);
      } catch {}
    }
    setInterval(renderDate, 60_000);
    connectAppBar();
    setUpInstalledApp();
  }

  // Installed desktop app (Edge "Install this site as an app")
  function setUpInstalledApp() {
    if (inOutlook || native || window.parent !== window) return; // not inside Outlook/Teams frames or the Android app
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("sw.js").catch(() => {});
    }
    // First launch as an app: open at a cozy to-do-list size instead of full screen
    if (matchMedia("(display-mode: standalone)").matches) {
      try {
        if (!localStorage.getItem("blossomSized")) {
          localStorage.setItem("blossomSized", "1");
          window.resizeTo(480, 760);
        }
      } catch {}
    }
  }

  // Running as an app in Outlook's left app bar (a Microsoft 365 personal tab)
  function connectAppBar() {
    const teams = window.microsoftTeams;
    if (!teams?.app) return;
    teams.app
      .initialize()
      .then(() => teams.app.getContext())
      .then((ctx) => {
        hostTheme = ctx?.app?.theme || null;
        applyTheme();
        teams.app.registerOnThemeChangeHandler((theme) => {
          hostTheme = theme;
          applyTheme();
        });
      })
      .catch(() => {}); // not inside Outlook/Teams, e.g. opened in a normal browser
  }

  if (window.Office && Office.onReady) {
    Office.onReady((info) => {
      inOutlook = info.host === Office.HostType.Outlook;
      init();
    });
  } else {
    init();
  }
})();
