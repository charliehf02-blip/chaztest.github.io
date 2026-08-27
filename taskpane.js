/* Halo Request Forms - task pane logic */

// CONFIGURE BEFORE DEPLOYMENT:
// Set this to your Halo ITSM intake mailbox so it's pre-filled in the To field.
// Leave as "" to skip pre-filling and let the user address the email themselves.
const HALO_INTAKE_EMAIL = ""; // e.g. "halo-intake@yourorg.com"

let currentType = "change"; // "change" | "closing"

Office.onReady(() => {
  document.getElementById("btnChange").addEventListener("click", () => switchType("change"));
  document.getElementById("btnClosing").addEventListener("click", () => switchType("closing"));
  document.getElementById("generateBtn").addEventListener("click", onGenerate);
});

function switchType(type) {
  currentType = type;
  const isChange = type === "change";

  document.getElementById("btnChange").classList.toggle("active", isChange);
  document.getElementById("btnClosing").classList.toggle("active", !isChange);
  document.getElementById("changeForm").style.display = isChange ? "block" : "none";
  document.getElementById("closingForm").style.display = isChange ? "none" : "block";

  setStatus("", "");
}

function onGenerate() {
  const form = currentType === "change"
    ? document.getElementById("changeForm")
    : document.getElementById("closingForm");

  if (!form.reportValidity()) {
    setStatus("Please fill in all required fields.", "error");
    return;
  }

  const fields = currentType === "change" ? collectChangeFields() : collectClosingFields();
  const subject = buildSubject(currentType, fields);
  const bodyHtml = buildBodyHtml(currentType, fields);

  const btn = document.getElementById("generateBtn");
  btn.disabled = true;
  setStatus("Inserting into email\u2026", "");

  const item = Office.context.mailbox.item;

  const tasks = [];

  tasks.push(new Promise((resolve, reject) => {
    item.subject.setAsync(subject, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  }));

  tasks.push(new Promise((resolve, reject) => {
    item.body.setAsync(
      bodyHtml,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(result.error);
      }
    );
  }));

  if (HALO_INTAKE_EMAIL) {
    tasks.push(new Promise((resolve, reject) => {
      item.to.setAsync(
        [{ displayName: "Halo ITSM", emailAddress: HALO_INTAKE_EMAIL }],
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(result.error);
        }
      );
    }));
  }

  Promise.all(tasks)
    .then(() => {
      setStatus("Inserted. Review the email, then send.", "success");
      btn.disabled = false;
    })
    .catch((err) => {
      console.error(err);
      setStatus("Something went wrong inserting the form. Please try again.", "error");
      btn.disabled = false;
    });
}

function collectChangeFields() {
  return [
    { label: "Request Type", value: "Change" },
    { label: "Urgency", value: val("c_urgency") },
    { label: "Matter Name", value: val("c_matterName") },
    { label: "Matter Number", value: val("c_matterNumber") },
    { label: "Client Number", value: val("c_clientNumber") },
    { label: "Client Name", value: val("c_clientName") },
    { label: "Information To Update", value: val("c_infoType") },
    { label: "Change Level", value: val("c_level") },
    { label: "Description", value: val("c_description"), multiline: true },
    { label: "3E EWalled", value: val("c_ewalled") },
  ];
}

function collectClosingFields() {
  return [
    { label: "Request Type", value: "File Closing" },
    { label: "Urgency", value: val("f_urgency") },
    { label: "Matter Name", value: val("f_matterName") },
    { label: "Matter Number", value: val("f_matterNumber") },
    { label: "Client Number", value: val("f_clientNumber") },
    { label: "Closure Reason", value: val("f_reason") },
    { label: "Final Billing Status", value: val("f_billing") },
  ];
}

function val(id) {
  return document.getElementById(id).value.trim();
}

function buildSubject(type, fields) {
  const matterName = fields.find(f => f.label === "Matter Name")?.value || "";
  const matterNumber = fields.find(f => f.label === "Matter Number")?.value || "";
  const label = type === "change" ? "Change Request" : "File Closing";
  return `[${label}] ${matterName} (${matterNumber})`;
}

// Renders as consistent "Label: value" lines, one per row, so Halo's
// mailbox field-mapping rules can parse them reliably. Confirm the exact
// format your Halo admin wants to map against before finalizing.
function buildBodyHtml(type, fields) {
  const title = type === "change" ? "Change Request" : "File Closing Request";
  const rows = fields.map(f => {
    const value = escapeHtml(f.value).replace(/\n/g, "<br/>");
    return `
      <tr>
        <td style="padding:4px 12px 4px 0;font-weight:600;color:#00355c;white-space:nowrap;vertical-align:top;">${escapeHtml(f.label)}:</td>
        <td style="padding:4px 0;color:#1a1a1a;">${value}</td>
      </tr>`;
  }).join("");

  return `
    <div style="font-family:Segoe UI, Arial, sans-serif; font-size:13px; color:#1a1a1a;">
      <h2 style="color:#00355c; font-size:15px; margin:0 0 12px 0;">${title}</h2>
      <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;">
        ${rows}
      </table>
    </div>`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function setStatus(message, kind) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = "status" + (kind ? " " + kind : "");
}
