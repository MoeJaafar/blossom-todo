// Blossom To-Do — Outlook task pane
// Tasks are saved in Outlook roaming settings (they follow your mailbox to
// every device). Opened outside Outlook, it falls back to localStorage so the
// page can be previewed in a normal browser.

(() => {
  const STORAGE_KEY = "blossomTodos";
  const $ = (id) => document.getElementById(id);

  let todos = [];
  let filter = "all";
  let inOutlook = false;

  // ---------- Storage ----------------------------------------------------

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
  };

  // ---------- Tasks ------------------------------------------------------

  function addTodo(text, email) {
    text = text.trim().slice(0, 300);
    if (!text) return;
    todos.unshift({
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      text,
      done: false,
      created: Date.now(),
      ...(email ? { email } : {}),
    });
    store.save();
    render();
  }

  function toggle(id, checkEl) {
    const t = todos.find((x) => x.id === id);
    if (!t) return;
    t.done = !t.done;
    store.save();
    if (t.done && checkEl) petals.burst(checkEl.getBoundingClientRect());
    render();
  }

  function remove(id) {
    todos = todos.filter((x) => x.id !== id);
    store.save();
    render();
  }

  function rename(id, text) {
    const t = todos.find((x) => x.id === id);
    text = text.trim().slice(0, 300);
    if (t && text && text !== t.text) {
      t.text = text;
      store.save();
    }
    render();
  }

  // ---------- Rendering --------------------------------------------------

  function render() {
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
      meta.textContent = new Date(t.created).toLocaleDateString(undefined, {
        month: "short",
        day: "numeric",
      });
    }
    body.appendChild(meta);

    const del = document.createElement("button");
    del.type = "button";
    del.className = "delete";
    del.textContent = "×";
    del.setAttribute("aria-label", "Delete task");
    del.addEventListener("click", () => remove(t.id));

    li.append(check, body, del);
    return li;
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
    todos = store.load();
    render();
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
      todos = todos.filter((t) => !t.done);
      store.save();
      render();
    });

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
    if (inOutlook || window.parent !== window) return; // not inside Outlook/Teams frames
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
