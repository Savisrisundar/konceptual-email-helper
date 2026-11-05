// The Outlook side panel: reads the open email, asks the local AI (ai.js) for a reply,
// and puts the reviewed draft into a reply. It never sends anything.

/* global Office, KonceptualAI */

const $ = (id) => document.getElementById(id);
const AUTOMATED = /no-?reply|do-?not-?reply|notifications?@|newsletter|mailer-daemon|postmaster|bounce/i;

let settings = { ...KonceptualAI.DEFAULTS, auto: true };
let running = null; // AbortController for the draft being written
let isCompose = false;

// ---------- Settings (saved with the mailbox, so they follow the user between computers) ----------

function loadSettings() {
  const rs = Office.context.roamingSettings;
  for (const key of ["about", "localModel", "localUrl", "auto"]) {
    const v = rs.get(key);
    if (v !== undefined && v !== null && v !== "") settings[key] = v;
  }
  $("about").value = settings.about;
  $("localModel").value = settings.localModel;
  $("localUrl").value = settings.localUrl;
  $("auto").checked = !!settings.auto;
}

function saveSettings() {
  const rs = Office.context.roamingSettings;
  settings.about = $("about").value.trim();
  settings.localModel = $("localModel").value.trim() || KonceptualAI.DEFAULTS.localModel;
  settings.localUrl = $("localUrl").value.trim() || KonceptualAI.DEFAULTS.localUrl;
  settings.auto = $("auto").checked;
  for (const key of ["about", "localModel", "localUrl", "auto"]) rs.set(key, settings[key]);
  rs.saveAsync(() => {
    $("saved").textContent = "Saved";
    setTimeout(() => ($("saved").textContent = ""), 2000);
  });
}

// ---------- Reading the email ----------

const getAsync = (fn) =>
  new Promise((resolve) => fn((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)));

async function readContext() {
  const item = Office.context.mailbox.item;
  const me = Office.context.mailbox.userProfile.displayName || "";
  const body = (await getAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb))) || "";

  if (!isCompose) {
    const from = item.from ? `${item.from.displayName || ""} <${item.from.emailAddress || ""}>`.trim() : "";
    const sentTo = (item.to || []).map((r) => r.displayName || r.emailAddress).filter(Boolean).join(", ");
    return { me, subject: item.subject || "", from, fromName: item.from?.displayName || "", sentTo, thread: body };
  }

  // Writing an email: a reply or forward has the earlier emails quoted in the body.
  const subject = (await getAsync((cb) => item.subject.getAsync(cb))) || "";
  const to = (await getAsync((cb) => item.to.getAsync(cb))) || [];
  const toText = to.map((r) => `${r.displayName || ""} <${r.emailAddress || ""}>`).join(", ");
  const isReply = /^(re|fw|fwd)\s*:/i.test(subject) || /\n\s*From:\s.+\n/i.test(body);
  return { me, subject, from: toText ? `(you are writing to) ${toText}` : "", thread: isReply ? body : "", typed: isReply ? "" : body.trim() };
}

function describeItem() {
  const item = Office.context.mailbox.item;
  if (!item) return "Open an email to get a suggested reply.";
  if (isCompose) return "Writing an email. Add notes below if you like, then press Write reply.";
  const who = item.from?.displayName || item.from?.emailAddress || "unknown sender";
  return `Reply to ${who}: ${item.subject || "(no subject)"}`;
}

// ---------- Writing the draft ----------

function setStatus(text, isError = false) {
  $("status").textContent = text;
  $("status").classList.toggle("error", isError);
  $("status").classList.remove("warn");
}

function setDraftButtons(enabled) {
  for (const id of ["reply", "reply-all", "copy", "insert", "copy2"]) $(id).disabled = !enabled;
}

async function write() {
  if (!Office.context.mailbox.item) return;
  if (running) running.abort();
  const controller = new AbortController();
  running = controller;

  $("write").disabled = true;
  $("stop").hidden = false;
  setDraftButtons(false);
  $("draft").value = "";
  setStatus("Writing…");
  const started = Date.now();

  try {
    const ctx = await readContext();
    ctx.notes = [$("notes").value.trim(), ctx.typed || ""].filter(Boolean).join("\n");
    if (!ctx.thread && !ctx.notes) {
      setStatus("Type a few words in the box above about what the email should say, then press Write reply.");
      return;
    }
    const text = await KonceptualAI.writeEmail(ctx, settings, (soFar) => ($("draft").value = soFar), controller.signal);
    if (running !== controller) return; // a newer draft has started
    $("draft").value = text;
    setDraftButtons(true);
    const secs = Math.round((Date.now() - started) / 1000);
    const warnings = KonceptualAI.checkDraft(text, ctx);
    setStatus(
      warnings.length
        ? `Draft ready (${secs} s). Please check: ${warnings.join("; ")}.`
        : `Draft ready (${secs} s). Read and edit before using it.`,
      false
    );
    $("status").classList.toggle("warn", warnings.length > 0);
  } catch (err) {
    if (err && err.name === "AbortError") return;
    setStatus(err instanceof KonceptualAI.AIError ? err.message : `Something went wrong: ${err?.message || err}`, true);
  } finally {
    if (running === controller) {
      running = null;
      $("write").disabled = false;
      $("stop").hidden = true;
    }
  }
}

function stop() {
  if (running) running.abort();
  running = null;
  $("write").disabled = false;
  $("stop").hidden = true;
  setStatus("Stopped.");
  if ($("draft").value.trim()) setDraftButtons(true);
}

// ---------- Using the draft ----------

function toHtml(text) {
  const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return text
    .split("\n")
    .map((line) => (line.trim() ? `<div>${esc(line)}</div>` : "<div><br></div>"))
    .join("");
}

// Opens Outlook's own reply window with the draft in it. It never sends: the user presses Send.
// The newer *Async version reports whether it worked; older Outlook only has the plain one.
function reply(all) {
  const text = $("draft").value.trim();
  const html = toHtml(text);
  const item = Office.context.mailbox.item;
  const opened = () =>
    setStatus("Reply opened in Outlook with the draft (it may be below the email: scroll down). Read it, then press Send there yourself.");
  const failed = async () => {
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      $("draft").select();
      document.execCommand("copy");
    }
    setStatus(`Outlook didn't open the reply, so the draft is copied. Click Outlook's own ${all ? "Reply all" : "Reply"} button, then press Ctrl+V in the reply.`, true);
  };
  try {
    const asyncFn = all ? item.displayReplyAllFormAsync : item.displayReplyFormAsync;
    if (typeof asyncFn === "function") {
      asyncFn.call(item, html, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? opened() : failed()));
    } else {
      if (all) item.displayReplyAllForm(html);
      else item.displayReplyForm(html);
      opened();
    }
  } catch {
    failed();
  }
}

function insert() {
  const item = Office.context.mailbox.item;
  const text = $("draft").value.trim();
  item.body.getTypeAsync((r) => {
    const html = r.value === Office.CoercionType.Html;
    const data = html ? toHtml(text) + "<div><br></div>" : text + "\n\n";
    item.body.prependAsync(data, { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text }, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) setStatus("Draft added to the top of your email. Read it, then press Send yourself.");
      else setStatus("Couldn't add it to the email. Press Copy, then paste it in yourself.", true);
    });
  });
}

async function copy() {
  const text = $("draft").value.trim();
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    $("draft").select();
    document.execCommand("copy");
  }
  setStatus("Copied. Paste it into your email with Ctrl+V.");
}

// ---------- Opening an email ----------

function onItem() {
  if (running) running.abort();
  running = null;
  const item = Office.context.mailbox.item;
  isCompose = !!item && typeof item.subject !== "string"; // in compose, subject is an object with getAsync
  $("read-actions").hidden = isCompose;
  $("compose-actions").hidden = !isCompose;
  $("about-item").textContent = describeItem();
  $("draft").value = "";
  $("notes").value = "";
  setDraftButtons(false);
  $("write").disabled = !item;
  $("stop").hidden = true;
  setStatus("");
  if (!item || isCompose || !settings.auto) return;
  if (AUTOMATED.test(item.from?.emailAddress || "")) {
    setStatus("Looks like an automated email, so no reply was written. Press Write reply if you want one.");
    return;
  }
  write();
}

// Shows when this copy was uploaded (publish.ps1 stamps ?v=yyyyMMddHHmm on the script links),
// so it's easy to see whether Outlook has loaded the latest version.
function showVersion() {
  const m = /[?&]v=(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})/.exec(document.querySelector('script[src*="taskpane.js"]')?.src || "");
  $("version").textContent = m ? `Version: ${+m[3]}/${+m[2]}/${m[1]} ${m[4]}:${m[5]}` : "Version: local test copy";
}

Office.onReady(() => {
  showVersion();
  loadSettings();
  $("write").addEventListener("click", write);
  $("stop").addEventListener("click", stop);
  $("reply").addEventListener("click", () => reply(false));
  $("reply-all").addEventListener("click", () => reply(true));
  $("insert").addEventListener("click", insert);
  $("copy").addEventListener("click", copy);
  $("copy2").addEventListener("click", copy);
  $("save").addEventListener("click", saveSettings);
  // When the panel is pinned, Outlook tells us each time a different email is selected.
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItem);
  } catch {
    // Older Outlook: the panel closes when another email is opened, so nothing to do.
  }
  onItem();
});
